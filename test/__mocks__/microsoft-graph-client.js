// Mock for @microsoft/microsoft-graph-client
const mockGraphClient = {
  api: jest.fn().mockReturnThis(),
  get: jest.fn().mockResolvedValue({ value: [] }),
  post: jest.fn().mockResolvedValue({ id: 'test-id' })
};

module.exports = {
  Client: {
    init: jest.fn().mockImplementation(() => mockGraphClient)
  },
  TokenCredentialAuthenticationProvider: jest.fn()
};
