import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createSmitheryUrl } from "@smithery/sdk/client/transport.js";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import dotenv from 'dotenv';

// Load environment variables
dotenv.config();

// Configuration for the client
const config = {
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  tenantId: process.env.TENANT_ID,
  redirectUri: process.env.REDIRECT_URI,
  apiKey: process.env.API_KEY,
  smitheryApiKey: process.env.SMITHERY_API_KEY
};

// Validate required configuration
const requiredFields = ['clientId', 'clientSecret', 'tenantId', 'redirectUri', 'apiKey', 'smitheryApiKey'];
for (const field of requiredFields) {
  if (!config[field]) {
    throw new Error(`Missing required configuration field: ${field}`);
  }
}

async function main() {
  try {
    // Create MCP client
    const client = new Client({
      name: "Test client",
      version: "1.0.0"
    });

    // Create transport
    const transport = new StreamableHTTPClientTransport();

    // Connect to server
    const serverUrl = createSmitheryUrl({
      apiKey: config.smitheryApiKey,
      hostname: 'localhost',
      port: 3000,
      path: '/mcp'
    });

    // Initialize connection with config
    await client.initialize(transport, serverUrl, { config });

    // List available tools
    const tools = await client.listTools();
    console.log('Available tools:', tools);

    // Keep the connection alive
    await new Promise(() => {});
  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
}

main();
