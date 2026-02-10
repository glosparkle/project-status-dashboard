# Project Status Dashboard

Web app that pulls Outlook emails, Teams channel messages, and flagged items to build a status summary and action items for each project.

## Setup

1. Create an Azure App Registration
   - Azure Portal > Azure Active Directory > App registrations > New registration
   - Name: `Project Status Dashboard`
   - Supported account types: choose the tenant you use (single tenant is fine)
   - Redirect URI: `http://localhost:5173`

2. Add permissions (Microsoft Graph, Delegated)
   - `User.Read`
   - `Mail.Read`
   - `Mail.ReadBasic`
   - `Chat.Read`
   - `ChannelMessage.Read.All`
   - `Team.ReadBasic.All`
   - `Group.Read.All`

3. Create `.env` from `.env.example`
   - `VITE_MSAL_CLIENT_ID`
   - `VITE_MSAL_TENANT_ID` (optional, leave blank for common)
   - `VITE_MSAL_REDIRECT_URI`

4. Install and run

```bash
npm install
npm run dev
```

## How it works

- The app scans the last 14 days of Outlook messages and Teams channel messages.
- It filters each project by sender and keyword.
- “Misc (Flagged)” shows any flagged email regardless of sender.
- Summaries and action items are generated locally as a first-pass draft.

## Customize projects

Edit `src/config.js` to change senders, keywords, and channel names.
