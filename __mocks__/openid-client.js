// Mock for openid-client
class MockAuthorizationCodePKCE {
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

const mockIssuer = {
  Client: {
    register: jest.fn().mockResolvedValue({
      client_id: 'test-client-id',
      client_secret: 'test-client-secret',
      redirect_uris: ['http://localhost:3000/auth/callback'],
      token_endpoint_auth_method: 'none',
      authorizationUrl: 'http://mock-auth-url.com',
      tokenUrl: 'http://mock-token-url.com',
      userinfoUrl: 'http://mock-userinfo-url.com'
    })
  }
};

mockIssuer.discover = jest.fn().mockResolvedValue(mockIssuer);

module.exports = {
  Issuer: {
    discover: jest.fn().mockResolvedValue(mockIssuer)
  },
  Strategy: jest.fn(),
  TokenSet: jest.fn(),
  custom: {
    httpClient: jest.fn()
  },
  generators: {
    state: jest.fn().mockReturnValue('test-state'),
    codeVerifier: jest.fn().mockReturnValue('test-code-verifier')
  },
  AuthorizationCodePKCE: MockAuthorizationCodePKCE
};
