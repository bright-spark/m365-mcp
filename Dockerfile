FROM node:18-alpine

WORKDIR /app

# Copy package.json and package-lock.json
COPY package*.json ./

# Install dependencies
RUN pnpm install

# Copy app source
COPY . .

# Create volume for persistent data
VOLUME /app/data

# Expose the server port
EXPOSE 3000

# Set environment variables (these will be overridden in production)
ENV NODE_ENV=production \
    PORT=3000

# Run the application
CMD ["node", "m365-mcp.js"]
