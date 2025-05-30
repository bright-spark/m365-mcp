import { app } from './src/server/index.js';

// Only start the server if this file is run directly
if (import.meta.url === `file://${process.argv[1]}`) {
  const port = process.env.PORT || 3000;
  const server = app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });

  // Handle graceful shutdown
  process.on('SIGTERM', () => {
    console.log('SIGTERM received. Shutting down gracefully');
    server.close(() => {
      console.log('Process terminated');
    });
  });
}

export { app };
