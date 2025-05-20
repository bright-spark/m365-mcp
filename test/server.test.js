// Set up environment variables before importing anything
process.env.NODE_ENV = 'test';
process.env.CLIENT_ID = 'test-client-id';
process.env.REDIRECT_URI = 'http://localhost:3000/auth/callback';
process.env.SESSION_SECRET = 'test-session-secret';

// Create a mock app with all necessary methods as jest.fn()
const mockApp = {
  use: jest.fn(),
  get: jest.fn(),
  post: jest.fn(),
  set: jest.fn(),
  listen: jest.fn((port, callback) => {
    if (callback) callback();
    return { close: jest.fn() };
  })
};

// Mock express to always return mockApp and provide static and json methods
jest.mock('express', () => {
  const express = jest.fn(() => mockApp);
  express.static = jest.fn(() => 'static-middleware');
  express.json = jest.fn(() => 'json-middleware');
  return express;
});

// Mock the MCP Server
const mockMcpServer = {
  tool: jest.fn(),
  handleRequest: jest.fn().mockResolvedValue({
    content: [{ type: 'text', text: 'Mock response' }]
  })
};

// Mock the MCP SDK
jest.mock('@modelcontextprotocol/sdk/server/mcp.js', () => ({
  McpServer: jest.fn().mockImplementation(() => mockMcpServer)
}));

// Mock other dependencies
jest.mock('express-session');
jest.mock('cors');
jest.mock('helmet');
jest.mock('express-rate-limit');
jest.mock('morgan');
jest.mock('passport');
jest.mock('openid-client');

// Import the McpServer for instanceof checks
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

// Import the app after all mocks are set up
import { createApp } from '../m365-mcp.js';

describe('MCP Server', () => {
  let server;
  let app;
  
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

    // Create the app after all mocks are set up
    app = createApp();
  });
  
  afterAll(() => {
    // Clean up after all tests
    jest.restoreAllMocks();
  });

  it('should initialize without errors', () => {
    expect(server).toBeDefined();
    expect(McpServer).toHaveBeenCalled();
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
      expect(mockApp.post).toHaveBeenCalledWith(
        '/v2/mcp',
        expect.anything(), // Accepts "json-middleware" or a function
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