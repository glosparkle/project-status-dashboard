import React, { useEffect, useMemo, useState } from "react";
import {
  ensureInitialized,
  fetchEmailItems,
  fetchTeamsChannelMatches,
  fetchChannelMessages,
  generateActionItems,
  generateSummary,
  getActiveAccount,
  signIn,
  signOut
} from "./graph.js";
import { appSettings, projects } from "./config.js";

const channelStorageKey = "channelSelections";

function loadStoredSelections() {
  try {
    const raw = localStorage.getItem(channelStorageKey);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveSelections(selections) {
  localStorage.setItem(channelStorageKey, JSON.stringify(selections));
}

export default function App() {
  const [account, setAccount] = useState(null);
  const [itemsByProject, setItemsByProject] = useState({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [channelMatches, setChannelMatches] = useState({});
  const [channelSelections, setChannelSelections] = useState(() => loadStoredSelections());

  useEffect(() => {
    let active = true;
    ensureInitialized()
      .then(() => {
        if (!active) return;
        setAccount(getActiveAccount());
      })
      .catch((err) => {
        if (!active) return;
        setError(err.message || "MSAL initialization failed");
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (channelSelections) saveSelections(channelSelections);
  }, [channelSelections]);

  const connected = Boolean(account);

  const handleSignIn = async () => {
    setError("");
    try {
      const acct = await signIn();
      setAccount(acct);
    } catch (err) {
      setError(err.message || "Sign-in failed");
    }
  };

  const handleSignOut = async () => {
    await signOut();
    setAccount(null);
    setItemsByProject({});
  };

  const handleDiscoverChannels = async (project) => {
    setError("");
    try {
      const matches = await fetchTeamsChannelMatches(project.teamsChannelName);
      setChannelMatches((prev) => ({ ...prev, [project.id]: matches }));
      if (matches.length === 1) {
        setChannelSelections((prev) => ({
          ...prev,
          [project.id]: matches[0]
        }));
      }
    } catch (err) {
      setError(err.message || "Channel discovery failed");
    }
  };

  const refresh = async () => {
    setLoading(true);
    setError("");

    try {
      const emailItems = await fetchEmailItems(projects, appSettings.lookbackDays);
      const grouped = projects.reduce((acc, project) => {
        acc[project.id] = emailItems.filter((item) => item.projectId === project.id);
        return acc;
      }, {});

      for (const project of projects) {
        if (!project.teamsChannelName) continue;
        const selection = channelSelections[project.id];
        if (!selection) continue;
        const channelItems = await fetchChannelMessages(
          selection.teamId,
          selection.channelId,
          appSettings.lookbackDays
        );
        grouped[project.id] = [...(grouped[project.id] || []), ...channelItems];
      }

      setItemsByProject(grouped);
    } catch (err) {
      setError(err.message || "Refresh failed");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (connected) {
      refresh();
    }
  }, [connected]);

  return (
    <div className="app">
      <header className="header">
        <div>
          <p className="eyebrow">Status Hub</p>
          <h1>Project Status + Actions</h1>
          <p className="subhead">Last {appSettings.lookbackDays} days from Outlook + Teams</p>
        </div>
        <div className="header-actions">
          {connected ? (
            <>
              <button className="secondary" onClick={refresh} disabled={loading}>
                {loading ? "Refreshing..." : "Refresh"}
              </button>
              <button className="ghost" onClick={handleSignOut}>
                Sign out
              </button>
            </>
          ) : (
            <button className="primary" onClick={handleSignIn}>
              Sign in with Microsoft
            </button>
          )}
        </div>
      </header>

      {!connected && (
        <section className="card hero">
          <h2>Connect to Microsoft 365</h2>
          <p>
            This app pulls Outlook emails, Teams channel messages, and flagged items to
            build project status summaries and action lists.
          </p>
        </section>
      )}

      {error && (
        <section className="card error">
          <strong>Something went wrong.</strong>
          <span>{error}</span>
        </section>
      )}

      <section className="grid">
        {projects.map((project) => {
          const items = itemsByProject[project.id] || [];
          const summary = generateSummary(items);
          const actions = generateActionItems(items);
          const matches = channelMatches[project.id] || [];
          const selection = channelSelections[project.id];

          return (
            <article key={project.id} className="card">
              <div className="card-header">
                <div>
                  <h2>{project.name}</h2>
                  <p className="meta">{items.length} items captured</p>
                </div>
                {project.teamsChannelName && (
                  <button
                    className="ghost"
                    onClick={() => handleDiscoverChannels(project)}
                    disabled={!connected}
                  >
                    Find Channel
                  </button>
                )}
              </div>

              {project.teamsChannelName && matches.length > 0 && (
                <div className="field">
                  <label>Teams channel match</label>
                  <select
                    value={selection ? selection.channelId : ""}
                    onChange={(event) => {
                      const chosen = matches.find(
                        (match) => match.channelId === event.target.value
                      );
                      if (chosen) {
                        setChannelSelections((prev) => ({
                          ...prev,
                          [project.id]: chosen
                        }));
                      }
                    }}
                  >
                    <option value="">Select a channel</option>
                    {matches.map((match) => (
                      <option key={match.channelId} value={match.channelId}>
                        {match.channelName} · {match.teamName}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="summary">
                <div>
                  <h3>Progress</h3>
                  <pre>{summary.progress}</pre>
                </div>
                <div>
                  <h3>Blockers</h3>
                  <pre>{summary.blockers}</pre>
                </div>
                <div>
                  <h3>Deadlines</h3>
                  <pre>{summary.deadlines}</pre>
                </div>
                <div>
                  <h3>Next Steps</h3>
                  <pre>{summary.nextSteps}</pre>
                </div>
              </div>

              <div className="actions">
                <h3>Action Items</h3>
                <ul>
                  {actions.map((action, index) => (
                    <li key={`${project.id}-action-${index}`}>{action}</li>
                  ))}
                </ul>
              </div>

              <div className="items">
                <h3>Recent Items</h3>
                {items.length === 0 ? (
                  <p className="meta">No matches yet.</p>
                ) : (
                  items
                    .slice(0, 8)
                    .map((item) => (
                      <div key={item.id} className="item">
                        <div>
                          <strong>{item.subject}</strong>
                          <p className="meta">
                            {item.source} · {item.sender} · {new Date(item.date).toLocaleString()}
                          </p>
                          <p>{item.preview}</p>
                        </div>
                        {item.url && (
                          <a href={item.url} target="_blank" rel="noreferrer">
                            Open
                          </a>
                        )}
                      </div>
                    ))
                )}
              </div>
            </article>
          );
        })}
      </section>
    </div>
  );
}
