# Microsoft 365 (M365) MCP

Welcome to the Microsoft 365 (M365) MCP project!

## Overview
This project implements an MCP (Multi-Cloud Platform) integration for Microsoft 365 (M365). It provides tools, scripts, and configurations to interact with Microsoft 365 services in a multi-cloud or hybrid environment.

## Features
- Integrates Microsoft 365 with MCP workflows
- Supports configuration via `.env` files
- Docker and docker-compose support for easy deployment
- Includes documentation and contribution guidelines

## Getting Started
1. **Clone the repository:**
   ```sh
   git clone https://github.com/your-org/m365-mcp.git
   cd m365-mcp
   ```
2. **Install dependencies:**
   ```sh
   npm install
   # or
   pnpm install
   ```
3. **Configure environment:**
   - Copy `.env.example` to `.env.local` and update credentials as needed.

4. **Run the project:**
   ```sh
   node m365-mcp.js
   ```
   Or use Docker:
   ```sh
   docker-compose up
   ```

## Documentation
- See the `docs/` directory for detailed documentation.
- Refer to `CONTRIBUTING.md` for contribution guidelines.
- Review `CODE_OF_CONDUCT.md` for community standards.

## License
This project is licensed under the terms of the LICENSE file.

## Contact
For questions, issues, or contributions, please open an issue or pull request on GitHub.

## MCP Integration

You can connect this server to any MCP-compatible client (such as Cursor) using the following configuration:

### Example mcp.json
```json
{
  "name": "Outlook MCP Server",
  "description": "Model Context Protocol server for Microsoft Outlook with browser-based authentication",
  "url": "http://localhost:3000/v2/mcp",
  "version": "1.0.0",
  "type": "http"
}
```

### How to Use
1. Save the above JSON as `mcp.json`.
2. In your MCP client (e.g., Cursor), use the "Add MCP" or "Import MCP" feature and select this file, or paste the URL directly.
3. Make sure your server is running (`pnpm dev` or `pnpm start`).
4. Authenticate via the browser if prompted.

For more details, see the [docs/API.md](docs/API.md) and [docs/SECURITY.md](docs/SECURITY.md).
