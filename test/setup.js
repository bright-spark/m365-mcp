// Mock the AuthorizationCode class before the app loads
class MockAuthorizationCode {
  constructor(config) {
    this.config = config;
    this.authorizationCode = {
      authorizeURL: jest.fn().mockReturnValue('http://mock-auth-url.com')
    };
  }
  
  createTokenRequest() {
    return Promise.resolve({
      access_token: 'test-access-token',
      refresh_token: 'test-refresh-token',
      expires_in: 3600
    });
  }
}

global.AuthorizationCode = MockAuthorizationCode;

// Mock other required globals
global.fetch = jest.fn();
global.Headers = jest.fn();
global.Request = jest.fn();
global.Response = jest.fn();

// Mock process.env
process.env.NODE_ENV = 'test';
process.env.CLIENT_ID = 'test-client-id';
process.env.REDIRECT_URI = 'http://localhost:3000/auth/callback';
process.env.SESSION_SECRET = 'test-session-secret';
