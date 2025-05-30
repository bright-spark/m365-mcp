import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { ClientSecretCredential } from "@azure/identity";

// Default retry options
const defaultRetryOptions = {
  maxRetries: 3,
  retryDelay: 1000, // 1 second
  maxRetryDelay: 5000, // 5 seconds
  retryableStatusCodes: [408, 429, 500, 502, 503, 504]
};

// Default timeout in milliseconds (30 seconds)
const DEFAULT_TIMEOUT = 30000;

/**
 * Creates a middleware function that implements retry logic
 */
function createRetryMiddleware(options = defaultRetryOptions) {
  return async (context, next) => {
    let attempts = 0;
    let lastError;

    while (attempts <= options.maxRetries) {
      try {
        // Add timeout to the request
        const timeout = new Promise((_, reject) => {
          setTimeout(() => reject(new Error('Request timeout')), DEFAULT_TIMEOUT);
        });

        // Race between the actual request and the timeout
        const response = await Promise.race([next(), timeout]);
        return response;
      } catch (error) {
        lastError = error;
        const statusCode = error.statusCode || error.code;

        // Check if we should retry
        if (!options.retryableStatusCodes.includes(statusCode) || attempts === options.maxRetries) {
          throw error;
        }

        // Calculate delay with exponential backoff
        const delay = Math.min(
          options.retryDelay * Math.pow(2, attempts),
          options.maxRetryDelay
        );

        // Wait before retrying
        await new Promise(resolve => setTimeout(resolve, delay));
        attempts++;
      }
    }

    throw lastError;
  };
}

/**
 * Creates a Microsoft Graph client with retry logic and timeout
 */
export function createGraphClient(config) {
  const credential = new ClientSecretCredential(
    config.tenantId,
    config.clientId,
    config.clientSecret
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });

  return Client.initWithMiddleware({
    authProvider: authProvider,
    middleware: [
      createRetryMiddleware()
    ]
  });
}
