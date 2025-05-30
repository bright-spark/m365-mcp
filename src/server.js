import express from 'express';
import { Server } from '@modelcontextprotocol/sdk/dist/esm/index.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/dist/esm/server/streamableHttp.js';
import { OutlookCalendarTool } from './tools/outlook-calendar.js';
import { OutlookEmailTool } from './tools/outlook-email.js';
import { OutlookContactsTool } from './tools/outlook-contacts.js';
import { OutlookTasksTool } from './tools/outlook-tasks.js';
import helmet from 'helmet';
import cors from 'cors';
import rateLimit from 'express-rate-limit';
import morgan from 'morgan';

// Configure rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100 // limit each IP to 100 requests per windowMs
});

// Configure CORS
const corsOptions = {
  origin: process.env.ALLOWED_ORIGINS ? process.env.ALLOWED_ORIGINS.split(',') : '*',
  methods: ['GET', 'POST', 'DELETE'],
  allowedHeaders: ['Content-Type', 'mcp-session-id'],
  maxAge: 86400 // 24 hours
};

/**
 * Create a new MCP server instance with tools
 */
export function createServer(config) {
  const server = new Server(
    {
      name: 'M365 MCP Server',
      version: '1.0.0',
      description: 'MCP server for Microsoft 365 integration'
    },
    {
      capabilities: {
        tools: {}, // Enable tools
        resources: { subscribe: true }, // Enable resources with subscriptions
        prompts: {}, // Enable prompts if needed
        logging: {}, // Enable logging
        completions: {} // Enable completions
      }
    }
  );

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

  return { server };
}

// Create Express app
const app = express();

// Apply security middleware
app.use(helmet());
app.use(cors(corsOptions));
app.use(limiter);
app.use(morgan('combined'));

// Limit JSON body size
app.use(express.json({ limit: '10mb' }));

// Handle MCP requests
app.post('/mcp', async (req, res) => {
  try {
    const transport = new StreamableHTTPServerTransport();
    const { server } = createServer(req.body.params?.config);

    // Connect transport to server
    await server.connect(transport);

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
  if (!sessionId) {
    res.status(400).send('Invalid or missing session ID');
    return;
  }
  
  // Handle the request
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive'
  });
  res.write('\n');
});

// Handle DELETE requests for session termination
app.delete('/mcp', async (req, res) => {
  const sessionId = req.headers['mcp-session-id'];
  if (!sessionId) {
    res.status(400).send('Invalid or missing session ID');
    return;
  }
  
  const { transport } = sessions.get(sessionId);
  await transport.handleRequest(req, res);
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({
    jsonrpc: '2.0',
    error: {
      code: -32000,
      message: 'Internal server error',
      data: process.env.NODE_ENV === 'development' ? err.message : 'An error occurred'
    },
    id: null
  });
});

// Start server
const PORT = process.env.PORT || 3000;
const server = app.listen(PORT, () => {
  console.log(`MCP server listening on port ${PORT}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing HTTP server');
  server.close(() => {
    console.log('HTTP server closed');
    process.exit(0);
  });
});

export { app };
