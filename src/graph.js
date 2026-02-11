import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, graphScopes } from "./config.js";
import { subDays } from "date-fns";

const msalInstance = new PublicClientApplication(msalConfig);
const initPromise = msalInstance.initialize();

export async function ensureInitialized() {
  await initPromise;
  return msalInstance;
}

export async function signIn() {
  await ensureInitialized();
  const loginRequest = { scopes: graphScopes };
  const response = await msalInstance.loginPopup(loginRequest);
  msalInstance.setActiveAccount(response.account);
  return response.account;
}

export async function signOut() {
  await ensureInitialized();
  const account = msalInstance.getActiveAccount();
  if (account) {
    await msalInstance.logoutPopup({ account });
  }
}

export function getActiveAccount() {
  const account = msalInstance.getActiveAccount();
  if (account) return account;
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    return accounts[0];
  }
  return null;
}

async function getToken() {
  await ensureInitialized();
  const account = getActiveAccount();
  if (!account) throw new Error("No active account");
  try {
    const result = await msalInstance.acquireTokenSilent({
      account,
      scopes: graphScopes
    });
    return result.accessToken;
  } catch {
    const result = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return result.accessToken;
  }
}

async function graphGet(url) {
  const token = await getToken();
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      "ConsistencyLevel": "eventual"
    }
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph error ${response.status}: ${text}`);
  }
  return response.json();
}

async function fetchPaged(url) {
  const items = [];
  let next = url;
  while (next) {
    const data = await graphGet(next);
    if (Array.isArray(data.value)) items.push(...data.value);
    next = data["@odata.nextLink"] || null;
  }
  return items;
}

function normalizeText(text) {
  return (text || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function matchesProject(message, project) {
  const text = `${message.subject || ""} ${message.bodyPreview || ""}`.toLowerCase();
  const sender = (message.from?.emailAddress?.name || "").toLowerCase();
  const senderAddress = (message.from?.emailAddress?.address || "").toLowerCase();

  const senderMatch = project.senders.some((s) => {
    const needle = s.toLowerCase();
    return sender.includes(needle) || senderAddress.includes(needle);
  });

  const keywordMatch = project.keywords.some((k) => text.includes(k.toLowerCase()));

  return senderMatch || keywordMatch;
}

export async function fetchEmailItems(projects, lookbackDays) {
  const start = subDays(new Date(), lookbackDays).toISOString();
  const select = [
    "id",
    "subject",
    "bodyPreview",
    "from",
    "receivedDateTime",
    "flag"
  ].join(",");
  const filter = `receivedDateTime ge ${start}`;
  const url = `https://graph.microsoft.com/v1.0/me/messages?$select=${select}&$filter=${encodeURIComponent(
    filter
  )}&$top=50`;

  const messages = await fetchPaged(url);

  return projects.flatMap((project) => {
    if (project.includeFlaggedOnly) {
      return messages
        .filter((message) => message.flag?.flagStatus === "flagged")
        .map((message) => ({
          id: message.id,
          source: "Email",
          projectId: project.id,
          sender: message.from?.emailAddress?.name || message.from?.emailAddress?.address || "Unknown",
          date: message.receivedDateTime,
          subject: message.subject || "(No subject)",
          preview: message.bodyPreview || "",
          url: `https://outlook.office.com/mail/inbox/id/${message.id}`
        }));
    }

    return messages
      .filter((message) => matchesProject(message, project))
      .map((message) => ({
        id: message.id,
        source: "Email",
        projectId: project.id,
        sender: message.from?.emailAddress?.name || message.from?.emailAddress?.address || "Unknown",
        date: message.receivedDateTime,
        subject: message.subject || "(No subject)",
        preview: message.bodyPreview || "",
        url: `https://outlook.office.com/mail/inbox/id/${message.id}`
      }));
  });
}

export async function fetchTeamsChannelMatches(channelName) {
  const teams = await fetchPaged("https://graph.microsoft.com/v1.0/me/joinedTeams?$select=id,displayName");
  const matches = [];

  for (const team of teams) {
    const channels = await fetchPaged(
      `https://graph.microsoft.com/v1.0/teams/${team.id}/channels?$select=id,displayName`
    );
    for (const channel of channels) {
      if (channel.displayName?.toLowerCase() === channelName.toLowerCase()) {
        matches.push({
          teamId: team.id,
          teamName: team.displayName,
          channelId: channel.id,
          channelName: channel.displayName
        });
      }
    }
  }

  return matches;
}

export async function fetchChannelMessages(teamId, channelId, lookbackDays) {
  const start = subDays(new Date(), lookbackDays).toISOString();
  const select = "id,from,createdDateTime,body";
  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages?$select=${select}&$top=50`;
  const messages = await fetchPaged(url);

  return messages
    .filter((message) => message.createdDateTime >= start)
    .map((message) => ({
      id: message.id,
      source: "Teams",
      sender: message.from?.user?.displayName || "Unknown",
      date: message.createdDateTime,
      subject: "Teams message",
      preview: normalizeText(message.body?.content || ""),
      url: null
    }));
}

export function generateSummary(items) {
  const sorted = [...items].sort((a, b) => new Date(b.date) - new Date(a.date));
  const recent = sorted.slice(0, 10);
  const highlights = recent.map((item) => `- ${item.sender}: ${item.subject}`);

  return {
    progress: highlights.length ? highlights.join("\n") : "- No recent updates found",
    blockers: "- None detected (manual review recommended)",
    deadlines: "- None detected (manual review recommended)",
    nextSteps: "- Review recent items and confirm owners"
  };
}

export function generateActionItems(items) {
  const text = items.map((item) => `${item.subject} ${item.preview}`.toLowerCase());
  const actionPhrases = ["need", "please", "action", "follow up", "blocker", "deadline", "due", "next"];
  const actions = [];

  items.forEach((item) => {
    const lower = `${item.subject} ${item.preview}`.toLowerCase();
    if (actionPhrases.some((phrase) => lower.includes(phrase))) {
      actions.push(`Follow up on: ${item.subject} (${item.sender})`);
    }
  });

  if (!actions.length) {
    return ["Review latest items for potential actions"];
  }

  return actions.slice(0, 10);
}

export { msalInstance };
