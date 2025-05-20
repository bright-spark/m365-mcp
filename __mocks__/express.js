const express = () => {
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
  
  return mockApp;
};

express.static = jest.fn();
express.json = jest.fn();

module.exports = express;
