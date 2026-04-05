# Tool Registry

This document outlines the Model Context Protocol (MCP) Tools exposed to the LLM. 

By exposing these exact tool descriptions and JSON schemas (`src/tools/outlookTools.js`), the AI Agent intrinsically understands how to interact with the underlying `graphService.js` without knowing the complex Microsoft Graph query syntax.

## Tools Overview

### 1. \`read_emails\`
- **Purpose:** Fetches the most recent emails from the Microsoft inbox chronologically.
- **Parameters:**
  - `top` (number): Specifies the maximum number of latest emails to retrieve. Defaults to 10.
- **Under the hood:** Maps to `/me/messages` in Microsoft Graph utilizing `$orderby=receivedDateTime DESC` and `$top`.

### 2. \`search_emails\`
- **Purpose:** Identifies specific emails globally inside the user's mailbox matching a user-provided keyword.
- **Parameters:**
  - `keyword` (string - required): The exact keyword, subject line, or email address to search for.
- **Under the hood:** Maps to `/me/messages` using the built-in Microsoft Graph `$search` capability.

### 3. \`send_email\`
- **Purpose:** Composes and fires a plain-text email from the authenticated user's account to an external recipient.
- **Parameters:**
  - `to` (string - required): The recipient email.
  - `subject` (string - required): Subject line.
  - `body` (string - required): Plain-text body of the email.
- **Under the hood:** Fires a `POST` request to `/me/sendMail` and requests tracking in the Sent Items folder (`saveToSentItems: "true"`).

### 4. \`get_calendar_events\`
- **Purpose:** Retrieves the authenticated user's upcoming calendar events.
- **Parameters:**
  - `daysLookAhead` (number): Number of days forward from the exact current time to scan.
- **Under the hood:** Avoids manual recurrence calculation by mapping to the `/me/calendarView` endpoint, extracting start dates and timezone details cleanly.

### 5. \`create_calendar_event\`
- **Purpose:** Creates a standard calendar event block on the user's primary calendar.
- **Parameters:**
  - `title` (string - required): Event Subject.
  - `startTime` (string - required): Start ISO UTC date string.
  - `endTime` (string - required): End ISO UTC date string.
- **Under the hood:** Executes a `POST` to `/me/events`. Automatically forces `UTC` mapping so that timestamps resolved by AI models parse identically without localized timezone drift.

---

## Tool Execution Flow

1. **LLM Evaluation:** AI Agent determines which tool is required to satisfy the user's intent. The agent formats a JSON payload matching an `inputSchema`.
2. **Action Router (`server.js`):** The LLM's `CallToolRequestSchema` is triggered. The server delegates the specific action to `outlookTools.handleToolCall()`.
3. **Microsoft Graph (`graphService.js`):** The Graph API is queried via the authenticated `graphClient.js` middleware.
4. **Resolution:** Returns a flat, stringified JSON string back perfectly into the MCP STDIO transport.
