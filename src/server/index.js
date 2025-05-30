import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import express from 'express';
import { randomUUID } from 'node:crypto';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';
import { OutlookCalendarTool } from '../tools/outlook-calendar.js';
import { OutlookEmailTool } from '../tools/outlook-email.js';
import { OutlookContactsTool } from '../tools/outlook-contacts.js';
import { OutlookTasksTool } from '../tools/outlook-tasks.js';

// Map to store transports and tools by session ID
const sessions = new Map();

/**
 * Create a new MCP server instance with tools
 */
function createServer(config) {
  const server = new McpServer({
    name: 'M365 MCP Server',
    version: '1.0.0',
    description: 'MCP server for Microsoft 365 integration'
  });

  // Initialize tools with configuration
  const calendarTool = new OutlookCalendarTool();
  const emailTool = new OutlookEmailTool();
  const contactsTool = new OutlookContactsTool();
  const tasksTool = new OutlookTasksTool();

  // Initialize tools with configuration
  calendarTool.initialize(config);
  emailTool.initialize(config);
  contactsTool.initialize(config);
  tasksTool.initialize(config);

  // Register tools
  server.registerTool(calendarTool);
  server.registerTool(emailTool);
  server.registerTool(contactsTool);
  server.registerTool(tasksTool);

  return server;
}

// Create Express app
const app = express();
app.use(express.json());

// Handle POST requests for client-to-server communication
app.post('/mcp', async (req, res) => {
  try {
    // Check for existing session ID
    const sessionId = req.headers['mcp-session-id'];
    let transport;

    if (sessionId && sessions.has(sessionId)) {
      // Reuse existing transport
      transport = sessions.get(sessionId).transport;
    } else if (!sessionId && isInitializeRequest(req.body)) {
      // New initialization request
      transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (newSessionId) => {
          // Create new server instance for this session
          const server = createServer(req.body.params?.config);
          // Store the transport and server by session ID
          sessions.set(newSessionId, { transport, server });
        }
      });

      // Clean up when session is closed
      transport.onclose = () => {
        if (transport.sessionId) {
          sessions.delete(transport.sessionId);
        }
      };

      // Connect transport to server
      const server = createServer(req.body.params?.config);
      await server.connect(transport);
    } else {
      // Invalid request
      res.status(400).json({
        jsonrpc: '2.0',
        error: {
          code: -32000,
          message: 'Bad Request: No valid session ID provided',
        },
        id: null,
      });
      return;
    }

    // Handle the request
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error('Error handling request:', error);
    res.status(500).json({
      jsonrpc: '2.0',
      error: {
        code: -32000,
        message: 'Internal server error',
        data: error.message
      },
      id: null
    });
  }
});

// Handle GET requests for server-to-client notifications via SSE
app.get('/mcp', async (req, res) => {
  const sessionId = req.headers['mcp-session-id'];
  if (!sessionId || !sessions.has(sessionId)) {
    res.status(400).send('Invalid or missing session ID');
    return;
  }
  
  const { transport } = sessions.get(sessionId);
  await transport.handleRequest(req, res);
});

// Handle DELETE requests for session termination
app.delete('/mcp', async (req, res) => {
  const sessionId = req.headers['mcp-session-id'];
  if (!sessionId || !sessions.has(sessionId)) {
    res.status(400).send('Invalid or missing session ID');
    return;
  }
  
  const { transport } = sessions.get(sessionId);
  await transport.handleRequest(req, res);
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`MCP server listening on port ${PORT}`);
});

export { app };
