export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_MSAL_CLIENT_ID || "",
    authority: import.meta.env.VITE_MSAL_TENANT_ID
      ? `https://login.microsoftonline.com/${import.meta.env.VITE_MSAL_TENANT_ID}`
      : "https://login.microsoftonline.com/common",
    redirectUri: import.meta.env.VITE_MSAL_REDIRECT_URI || window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

export const graphScopes = [
  "User.Read",
  "Mail.Read",
  "Mail.ReadBasic",
  "Chat.Read",
  "ChannelMessage.Read.All",
  "Team.ReadBasic.All",
  "Group.Read.All"
];

export const appSettings = {
  lookbackDays: 14
};

export const projects = [
  {
    id: "mobile-credentials",
    name: "Mobile Credentials",
    keywords: ["mobile credentials", "digital badge"],
    senders: [
      "blake erickson",
      "brent harris",
      "james lowe",
      "carol westaway",
      "george cosare"
    ],
    teamsChannelName: "mobileCredentials"
  },
  {
    id: "ontic",
    name: "Ontic",
    keywords: ["ontic", "ontic api"],
    senders: [
      "phillip stix",
      "jack wiltbank",
      "george cosare",
      "sanah wong"
    ],
    teamsChannelName: "Ontic API"
  },
  {
    id: "misc-flagged",
    name: "Misc (Flagged)",
    keywords: [],
    senders: [],
    includeFlaggedOnly: true
  }
];
