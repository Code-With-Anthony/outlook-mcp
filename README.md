# Outlook MCP Server

This repository contains a Model Context Protocol (MCP) Server for Microsoft Outlook. It effectively bridges the gap between Microsoft Graph API and autonomous AI Agents (like Cline or Claude Desktop), allowing the AI to safely perform operations in your Outlook inbox and calendar.

## 🏗️ Architecture
This server follows a dual-application topology:
- **Express HTTP Server:** Runs persistently on `localhost:3000` waiting for the end-user to approve the OAuth 2.0 Microsoft Authentication flow.
- **MCP Server:** Runs locally over `stdio` strictly waiting to receive Tool Call JSON-RPC schemas from any connected MCP-compliant AI Agent.

---

## 🚀 Setup & Installation

### 1. Registering the Microsoft App
Before you start, you must register a new application within the [Azure Portal](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade):
1. **Name:** `Outlook MCP Server`
2. **Account Type:** Inter-org + Personal Accounts (Required if using standard outlook.com/hotmail.com).
3. **Redirect URI:** Select **Web** and add `http://localhost:3000/auth/callback`.
4. Create a **Client Secret** (save it temporarily).
5. Grant the following **Delegated** Graph API Permissions:
   - `User.Read`
   - `Mail.Read`
   - `Mail.Send`
   - `Calendars.Read`
   - `Calendars.ReadWrite`

### 2. Local Configuration
Clone this repo and configure your `.env` file based on your Azure keys.

```bash
npm install
```

Create exactly `.env` in the root:
```env
AZURE_CLIENT_ID="your_client_id_here"
AZURE_CLIENT_SECRET="your_client_secret_here"
AZURE_TENANT_ID="common" # usually "common" for outlook.com personal accounts
REDIRECT_URI="http://localhost:3000/auth/callback"
PORT=3000
```

---

## 🧪 Testing & Usage (Connecting to VS Code / Cline)

1. **Start the local server instance:**
   ```bash
   node src/server.js
   ```
2. Navigate your web browser to `http://localhost:3000/login`.
3. Sign into Microsoft and grant permissions. The system will hold your token actively in memory and automatically request silent refresh mechanisms going forward!
4. **Link to VS Code:** In VS Code, open the Cline Extension's MCP Settings (`cline_mcp.json`) and append the local executing environment:

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "node",
      "args": ["C:/absolute/path/to/repo/src/server.js"],
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

*Note: Replace the path above with your absolute system path to `src/server.js`.*
Once configured, you can ask Cline things like:
> `"Hey Cline, show me the 3 latest emails I received."`
> `"What meetings do I have tomorrow? Create one with John at 3 PM UTC."`

---

## 👨‍💻 Developer Guide: Adding New Capabilities

If you are expanding this codebase to support new Microsoft endpoints (e.g. OneDrive or ToDo), you MUST follow standard protocol implementation steps:

1. **Service Layer (`src/services/graphService.js`):** Draft your raw query mapping back to `graphClient.js` natively.
2. **Tool Definition (`src/tools/outlookTools.js`):** Expose the JSON schema describing input requirements within the `TOOLS` registry explicitly. Register your function string block onto the `handleToolCall` proxy payload router. 
3. **Important Limitation:** Never use `console.log()` across the codebase. As MCP establishes connectivity via Stdout streams, standard logging fatally manipulates JSON payloads and crashes AI connectivity. *Always use `console.error()` for diagnostic debugging.*
