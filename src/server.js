require('dotenv').config();
const express = require('express');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { CallToolRequestSchema, ListToolsRequestSchema } = require('@modelcontextprotocol/sdk/types.js');

const authRouter = require('./routes/auth');
const { TOOLS, handleToolCall } = require('./tools/outlookTools');

// ==========================================
// 1. EXPRESS HTTP SERVER (OAuth Flow)
// ==========================================
const app = express();
const PORT = process.env.PORT || 3000;

app.use('/', authRouter);

app.listen(PORT, () => {
    // IMPORTANT: We MUST use console.error instead of console.log
    // console.log writes to stdout, which corrupts the MCP stdio transport protocol!
    console.error(`🔒 Auth server running on http://localhost:${PORT}/login`);
    console.error("👉 Please navigate to this URL to authenticate your MCP Server with Microsoft.");
});

// ==========================================
// 2. MODEL CONTEXT PROTOCOL (MCP) SERVER
// ==========================================
const mcpServer = new Server(
    {
        name: "outlook-mcp-server",
        version: "1.0.0"
    },
    {
        capabilities: {
            tools: {}
        }
    }
);

// A. Register the tools schemas so LLMs know what functions exist
mcpServer.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: TOOLS
    };
});

// B. Execute the actual function when the LLM makes a tool call request
mcpServer.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    
    console.error(`[MCP] Tool invoked by AI Agent: ${name}`);

    try {
        const result = await handleToolCall(name, args);
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(result, null, 2) // Return JSON stringified result
                }
            ],
            isError: false
        };
    } catch (error) {
        console.error(`[MCP] Tool execution failed: ${error.message}`);
        return {
            content: [
                {
                    type: "text",
                    text: `Error executing tool '${name}': ${error.message}`
                }
            ],
            isError: true // Signals the AI Agent that the task failed so it can fix the input and retry
        };
    }
});

// C. Connect the MCP Server to Standard I/O
async function startMcpServer() {
    const transport = new StdioServerTransport();
    await mcpServer.connect(transport);
    console.error("✅ Outlook MCP Server is successfully listening on stdio.");
    console.error("🤖 Ready to receive commands from Cline, Claude Desktop, or other agents!");
}

startMcpServer().catch(err => {
    console.error("❌ Fatal error starting MCP server:", err);
    process.exit(1);
});
