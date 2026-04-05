# Authentication Architecture

This document describes how the Outlook MCP Server handles Microsoft Authentication. The system uses the industry-standard **OAuth 2.0 Authorization Code Flow** coupled with Microsoft's official Server-Side Identity library (**MSAL Node**) to securely authenticate users without handling sensitive passwords.

## Core Principles

1. **Passwordless on our end**: The MCP server never intercepts or sees the user's Microsoft password.
2. **Persistent Caching**: To prevent requiring developers to log in on every single server restart, access tokens and long-lived refresh tokens are seamlessly serialized to a secure, local system file.
3. **Silent Token Acquisition**: The API client handles token expiration automatically via MSAL's background refresh mechanism.

---

## The OAuth 2.0 Flow

| Step | Action | Description |
|---|---|---|
| **1. Authorization** | User navigates to `/login` | The local Express server redirects the user to Microsoft's secure login portal. |
| **2. Auth Code** | Microsoft Redirects | Upon success, Microsoft redirects the user to `/auth/callback` with a temporary, one-time-use **Authorization Code**. |
| **3. Token Exchange** | Server hits Microsoft API | Our MSAL client uses this Auth Code to request an **Access Token** (valid for 1 hour) and a **Refresh Token** (valid for months). |
| **4. Persistence** | Cache written to Disk | The `authConfig.js` cache plugin serializes the MSAL token cache into `src/config/msalTokenCache.json`. |
| **5. Session Resumption** | Server Restarts | On startup, MSAL deserializes `msalTokenCache.json` and silently issues new Access Tokens as needed. |

---

## File Responsibilities

### \`src/config/authConfig.js\`
- Initializes the `ConfidentialClientApplication` using the Azure Application Client ID and Secret (`dotenv`).
- Requests the exact Scopes required by the Graph API (`Mail.Read`, `Mail.Send`, `Calendars.ReadWrite`, `offline_access`). *Note: `offline_access` is required to yield a Refresh Token.*
- Implements the `cachePlugin` to intercept MSAL token reads/writes and persist them cleanly to disk.

### \`src/config/tokenStore.js\`
- Acts as a local heuristic validator. By extracting the Account payload from the local JSON cache, the tool can verify if a user is currently authenticated without actively attempting a failing Microsoft Graph network call.

### \`src/routes/auth.js\`
- A minimal Express router that exposes the endpoints needed to intercept Microsoft's OAuth web-hook callbacks.

### \`msalTokenCache.json\`
- **[DANGER]** This file contains live Authentication tokens. It is strictly included in `.gitignore` to prevent secret leaks to version control. If an attacker receives this file, they can impersonate the authorized user.
