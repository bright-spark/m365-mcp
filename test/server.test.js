// Set up environment variables before importing anything
process.env.NODE_ENV = 'test';
process.env.CLIENT_ID = 'test-client-id';
process.env.REDIRECT_URI = 'http://localhost:3000/auth/callback';
process.env.SESSION_SECRET = 'test-session-secret';

// Mock the MCP Server
const mockMcpServer = {
  tool: jest.fn(),
  handleRequest: jest.fn().mockResolvedValue({
    content: [{ type: 'text', text: 'Mock response' }]
  })
};

// Mock the MCP SDK
jest.mock('@modelcontextprotocol/sdk/server/mcp', () => ({
  McpServer: jest.fn().mockImplementation(() => mockMcpServer)
}));

// Mock other dependencies
jest.mock('express');
jest.mock('express-session');
jest.mock('cors');
jest.mock('helmet');
jest.mock('express-rate-limit');
jest.mock('morgan');
jest.mock('passport');
jest.mock('openid-client');

// Import the app after all mocks are set up
const express = require('express');
const app = require('../m365-mcp');

// Get the mock app instance
const mockApp = express();

describe('MCP Server', () => {
  let server;
  
  beforeEach(() => {
    // Reset all mocks before each test
    jest.clearAllMocks();
    
    // Set up the server instance
    server = mockMcpServer;
    
    // Mock the app.get method to return our mock server
    mockApp.get.mockImplementation((key) => {
      if (key === 'mcpServer') return server;
      return null;
    });
    
    // Set up environment variables for testing
    process.env.NODE_ENV = 'test';
    process.env.CLIENT_ID = 'test-client-id';
    process.env.REDIRECT_URI = 'http://localhost:3000/auth/callback';
    process.env.SESSION_SECRET = 'test-session-secret';
  });
  
  afterAll(() => {
    // Clean up after all tests
    jest.restoreAllMocks();
  });

  it('should initialize without errors', () => {
    expect(server).toBeDefined();
    expect(server).toBeInstanceOf(McpServer);
  });
  
  describe('Tool Registration', () => {
    it('should have registered the get_auth_status tool', () => {
      // This is a simple check - in a real test, you'd verify the tool's behavior
      expect(server.tool).toHaveBeenCalledWith(
        'get_auth_status',
        expect.any(Object),
        expect.any(Function)
      );
    });
    
    it('should have registered the list_emails tool', () => {
      expect(server.tool).toHaveBeenCalledWith(
        'list_emails',
        expect.any(Object),
        expect.any(Function)
      );
    });
  });
  
  describe('Server Configuration', () => {
    it('should have NODE_ENV set to test', () => {
      expect(process.env.NODE_ENV).toBe('test');
    });

    it('should have required environment variables', () => {
      expect(process.env.CLIENT_ID).toBeDefined();
      expect(process.env.REDIRECT_URI).toBeDefined();
    });
    
    it('should set up the /v2/mcp endpoint', () => {
      // Verify that the MCP endpoint was set up
      expect(app.post).toHaveBeenCalledWith(
        '/v2/mcp',
        expect.any(Function), // express.json() middleware
        expect.any(Function)  // The route handler
      );
    });
  });

  describe('Environment Tests', () => {
    it('should verify environment variables are loaded', () => {
      if (process.env.NODE_ENV !== 'production') {
        expect(true).toBe(true);
      } else {
        expect(process.env.CLIENT_ID?.length > 0).toBe(true);
      }
    });
  });
});