// Mock Express app
const express = function() {
  const app = function(req, res, next) { next(); };
  
  // Add methods
  app.use = jest.fn().mockReturnThis();
  app.get = jest.fn().mockReturnThis();
  app.post = jest.fn().mockReturnThis();
  app.set = jest.fn().mockReturnThis();
  app.listen = jest.fn((port, callback) => {
    if (callback) callback();
    return { close: jest.fn() };
  });
  
  return app;
};

// Add static method
express.static = jest.fn();

// Add json middleware
express.json = jest.fn();

module.exports = express;
