import express from 'express';
import session from 'express-session';
import cors from 'cors';
import bodyParser from 'body-parser';
import helmet from 'helmet';
import path from 'path';
import winston from 'winston';
import { Client } from '@microsoft/microsoft-graph-client';
import simpleOauth2 from 'simple-oauth2';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import crypto from 'crypto';
import fs from 'fs';
import dotenv from 'dotenv';
import { body, validationResult } from 'express-validator';
import { fileURLToPath } from 'url';

// Load environment variables
dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configuration
const config = {
  clientId: process.env.CLIENT_ID || 'your-client-id',
  redirectUri: process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback',
  port: process.env.PORT || 3000,
  sessionSecret: process.env.SESSION_SECRET || crypto.randomBytes(64).toString('hex')
};

// Microsoft Graph authentication settings
const authConfig = {
  authorizeHost: 'https://login.microsoftonline.com/common',
  tokenHost: 'https://login.microsoftonline.com/common',
  authorizePath: '/oauth2/v2.0/authorize',
  tokenPath: '/oauth2/v2.0/token'
};

// Required Microsoft Graph scopes
const scopes = [
  'offline_access',
  'openid',
  'profile',
  'User.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite'
];

// Store user tokens
// TODO: Use a persistent and secure store for tokens in production (e.g., Redis, encrypted DB)
const userTokens = {};

// Simple custom rate limiter middleware (per IP, 100 requests per 15 minutes)
const rateLimitWindowMs = 15 * 60 * 1000; // 15 minutes
const rateLimitMax = 100;
const ipRateLimits = new Map();

function customRateLimiter(req, res, next) {
  const now = Date.now();
  const ip = req.ip;
  let entry = ipRateLimits.get(ip);
  if (!entry || now - entry.start > rateLimitWindowMs) {
    entry = { count: 1, start: now };
    ipRateLimits.set(ip, entry);
  } else {
    entry.count++;
  }
  if (entry.count > rateLimitMax) {
    res.status(429).json({ error: 'Too many requests, please try again later.' });
    return;
  }
  next();
}

// Wrap all initialization code inside createApp
function createApp() {
  // Initialize Express app
  const app = express();

  // Security headers
  app.use(helmet());

  // Rate limiting
  app.use(customRateLimiter);

  app.use(cors({
    origin: 'http://localhost:3000',
    credentials: true
  }));
  app.use(bodyParser.json());
  app.use(session({
    secret: config.sessionSecret,
    resave: false,
    saveUninitialized: true,
    cookie: { secure: process.env.NODE_ENV === 'production' }
  }));

  // Serve static files
  app.use(express.static(path.join(__dirname, 'public')));

  // Logger setup
  const logger = winston.createLogger({
    level: 'info',
    format: winston.format.combine(
      winston.format.timestamp(),
      winston.format.json()
    ),
    transports: [
      new winston.transports.File({ filename: 'error.log', level: 'error' }),
      new winston.transports.File({ filename: 'combined.log' })
    ]
  });

  if (process.env.NODE_ENV !== 'production') {
    logger.add(new winston.transports.Console({
      format: winston.format.simple()
    }));
  }

  // Create a simple home page with login link
  app.get('/', (req, res) => {
    const isLoggedIn = !!(req.session.userId && userTokens[req.session.userId]);
    
    res.send(`
      <html>
        <head>
          <title>Outlook MCP Server</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              max-width: 800px;
              margin: 0 auto;
              padding: 20px;
            }
            .header {
              text-align: center;
              margin-bottom: 30px;
            }
            .container {
              border: 1px solid #ddd;
              border-radius: 8px;
              padding: 20px;
              margin-bottom: 20px;
            }
            .button {
              background-color: #0078d4;
              color: white;
              border: none;
              padding: 10px 20px;
              border-radius: 4px;
              text-decoration: none;
              display: inline-block;
              margin-top: 10px;
            }
            .status {
              margin-top: 20px;
              padding: 10px;
              border-radius: 4px;
            }
            .logged-in {
              background-color: #dff0d8;
              color: #3c763d;
            }
            .logged-out {
              background-color: #f2dede;
              color: #a94442;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <h1>Microsoft Outlook MCP Server</h1>
            <p>Model Context Protocol server for Outlook integration</p>
          </div>
          
          <div class="container">
            <h2>Server Status</h2>
            <div class="status ${isLoggedIn ? 'logged-in' : 'logged-out'}">
              <p><strong>Authentication Status:</strong> ${isLoggedIn ? 'Logged in to Microsoft Account' : 'Not logged in'}</p>
            </div>
            <p>
              ${isLoggedIn ? 
                '<a href="/auth/logout" class="button">Logout</a>' : 
                '<a href="/auth/login" class="button">Login with Microsoft</a>'}
            </p>
          </div>
          
          <div class="container">
            <h2>MCP Server Information</h2>
            <p>Endpoint URL: <code>http://localhost:${config.port}/v2/mcp</code></p>
            <p>Status: ${isLoggedIn ? 'Ready to accept MCP requests' : 'Please login first'}</p>
          </div>
        </body>
      </html>
    `);
  });
  
  // Define MCP tools with browser-based authentication
  const mcpTools = [
    // Authentication tools for checking login status
    {
      name: 'get_auth_status',
      description: 'Check if the user is authenticated with Microsoft Outlook',
      parameters: {
        type: 'object',
        properties: {}
      },
      handler: async ({ user_id }) => {
        try {
          const isAuthenticated = !!(user_id && userTokens[user_id]);
          
          if (!isAuthenticated) {
            return {
              authenticated: false,
              login_url: `http://localhost:${config.port}/auth/login`
            };
          }
          
          // Check if we have a valid token
          try {
            await getValidAccessToken(user_id);
            return { authenticated: true };
          } catch (error) {
            return {
              authenticated: false,
              login_url: `http://localhost:${config.port}/auth/login`,
              error: error.message
            };
          }
        } catch (error) {
          logger.error('Error getting auth status:', error);
          throw new Error(`Failed to get authentication status: ${error.message}`);
        }
      }
    },
  
    // Mail operations
    {
      name: 'list_emails',
      description: 'List recent emails from the Outlook inbox',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          query: { 
            type: 'string', 
            description: 'Search query to filter emails'
          },
          maxResults: { 
            type: 'number', 
            description: 'Maximum number of emails to return (default: 10)'
          }
        },
        required: ['user_id']
      },
      handler: async ({ user_id, query, maxResults = 10 }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          let endpoint = '/me/messages';
          if (query) {
            endpoint = `/me/messages?$filter=contains(subject,'${query}')`;
          }
          
          const response = await graphClient
            .api(endpoint)
            .top(maxResults)
            .select('id,subject,bodyPreview,from,receivedDateTime')
            .orderBy('receivedDateTime DESC')
            .get();
            
          return response.value;
        } catch (error) {
          logger.error('Error listing emails:', error);
          throw new Error(`Failed to list emails: ${error.message}`);
        }
      }
    },
    
    {
      name: 'search_emails',
      description: 'Search emails with advanced query',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          query: { 
            type: 'string', 
            description: 'Outlook search query (e.g., "from:example@gmail.com has:attachment")',
            required: true
          },
          maxResults: { 
            type: 'number', 
            description: 'Maximum number of emails to return (default: 10)'
          }
        },
        required: ['user_id', 'query']
      },
      handler: async ({ user_id, query, maxResults = 10 }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          const response = await graphClient
            .api('/me/messages')
            .search(query)
            .top(maxResults)
            .select('id,subject,bodyPreview,from,receivedDateTime,hasAttachments')
            .orderBy('receivedDateTime DESC')
            .get();
            
          return response.value;
        } catch (error) {
          logger.error('Error searching emails:', error);
          throw new Error(`Failed to search emails: ${error.message}`);
        }
      }
    },
    
    {
      name: 'get_email',
      description: 'Get a specific email by ID',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          id: { 
            type: 'string', 
            description: 'Email ID',
            required: true
          },
          format: {
            type: 'string',
            description: 'Format to retrieve the email (html or text)',
            enum: ['html', 'text'],
            default: 'html'
          }
        },
        required: ['user_id', 'id']
      },
      handler: async ({ user_id, id, format = 'html' }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          const response = await graphClient
            .api(`/me/messages/${id}`)
            .select('id,subject,body,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments')
            .get();
          
          // Return formatted body based on format parameter
          if (format === 'text' && response.body.content) {
            // Simple HTML to text conversion
            response.body.content = response.body.content
              .replace(/<[^>]*>/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();
          }
            
          return response;
        } catch (error) {
          logger.error('Error getting email:', error);
          throw new Error(`Failed to get email: ${error.message}`);
        }
      }
    },
    
    {
      name: 'send_email',
      description: 'Send a new email',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          to: { 
            type: 'string', 
            description: 'Recipient email address',
            required: true
          },
          subject: { 
            type: 'string', 
            description: 'Email subject',
            required: true
          },
          body: { 
            type: 'string', 
            description: 'Email body (can include HTML)',
            required: true
          },
          cc: { 
            type: 'string', 
            description: 'CC recipients (comma-separated)'
          },
          bcc: { 
            type: 'string', 
            description: 'BCC recipients (comma-separated)'
          },
          isHtml: {
            type: 'boolean',
            description: 'Whether the body is HTML',
            default: true
          }
        },
        required: ['user_id', 'to', 'subject', 'body']
      },
      handler: async ({ user_id, to, subject, body, cc, bcc, isHtml = true }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          const toRecipients = to.split(',').map(email => ({
            emailAddress: {
              address: email.trim()
            }
          }));
          
          const ccRecipients = cc ? cc.split(',').map(email => ({
            emailAddress: {
              address: email.trim()
            }
          })) : [];
          
          const bccRecipients = bcc ? bcc.split(',').map(email => ({
            emailAddress: {
              address: email.trim()
            }
          })) : [];
          
          const message = {
            subject,
            body: {
              contentType: isHtml ? 'HTML' : 'Text',
              content: body
            },
            toRecipients,
            ccRecipients,
            bccRecipients
          };
          
          await graphClient
            .api('/me/sendMail')
            .post({
              message,
              saveToSentItems: true
            });
            
          return { success: true, message: 'Email sent successfully' };
        } catch (error) {
          logger.error('Error sending email:', error);
          throw new Error(`Failed to send email: ${error.message}`);
        }
      }
    },
    
    // Calendar operations
    {
      name: 'list_events',
      description: 'List upcoming calendar events',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          timeMin: { 
            type: 'string', 
            description: 'Start time in ISO format (default: now)'
          },
          timeMax: { 
            type: 'string', 
            description: 'End time in ISO format'
          },
          maxResults: { 
            type: 'number', 
            description: 'Maximum number of events to return (default: 10)'
          }
        },
        required: ['user_id']
      },
      handler: async ({ user_id, timeMin, timeMax, maxResults = 10 }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          let queryOptions = '';
          
          if (timeMin && timeMax) {
            const encodedTimeMin = encodeURIComponent(timeMin);
            const encodedTimeMax = encodeURIComponent(timeMax);
            queryOptions = `?$filter=start/dateTime ge '${encodedTimeMin}' and end/dateTime le '${encodedTimeMax}'`;
          } else if (timeMin) {
            const encodedTimeMin = encodeURIComponent(timeMin);
            queryOptions = `?$filter=start/dateTime ge '${encodedTimeMin}'`;
          } else if (timeMax) {
            const encodedTimeMax = encodeURIComponent(timeMax);
            queryOptions = `?$filter=end/dateTime le '${encodedTimeMax}'`;
          }
          
          const response = await graphClient
            .api(`/me/events${queryOptions}`)
            .top(maxResults)
            .select('id,subject,bodyPreview,start,end,location,organizer,attendees')
            .orderBy('start/dateTime')
            .get();
            
          return response.value;
        } catch (error) {
          logger.error('Error listing events:', error);
          throw new Error(`Failed to list events: ${error.message}`);
        }
      }
    },
    
    {
      name: 'create_event',
      description: 'Create a new calendar event',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          subject: { 
            type: 'string', 
            description: 'Event title',
            required: true
          },
          start: { 
            type: 'string', 
            description: 'Start time in ISO format',
            required: true
          },
          end: { 
            type: 'string', 
            description: 'End time in ISO format',
            required: true
          },
          location: { 
            type: 'string', 
            description: 'Event location'
          },
          description: { 
            type: 'string', 
            description: 'Event description'
          },
          attendees: { 
            type: 'array', 
            items: { type: 'string' },
            description: 'List of attendee email addresses'
          },
          isOnlineMeeting: {
            type: 'boolean',
            description: 'Whether this is an online meeting',
            default: false
          }
        },
        required: ['user_id', 'subject', 'start', 'end']
      },
      handler: async ({ user_id, subject, start, end, location, description, attendees, isOnlineMeeting = false }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          const eventAttendees = attendees ? attendees.map(email => ({
            emailAddress: {
              address: email
            },
            type: 'required'
          })) : [];
          
          const event = {
            subject,
            start: {
              dateTime: start,
              timeZone: 'UTC'
            },
            end: {
              dateTime: end,
              timeZone: 'UTC'
            },
            body: {
              contentType: 'HTML',
              content: description || ''
            },
            isOnlineMeeting,
            onlineMeetingProvider: isOnlineMeeting ? 'teamsForBusiness' : null
          };
          
          if (location) {
            event.location = {
              displayName: location
            };
          }
          
          if (eventAttendees.length > 0) {
            event.attendees = eventAttendees;
          }
          
          const response = await graphClient
            .api('/me/events')
            .post(event);
            
          return response;
        } catch (error) {
          logger.error('Error creating event:', error);
          throw new Error(`Failed to create event: ${error.message}`);
        }
      }
    },
    
    // Contacts operations
    {
      name: 'list_contacts',
      description: 'List contacts from Outlook',
      parameters: {
        type: 'object',
        properties: {
          user_id: { 
            type: 'string', 
            description: 'User ID from authentication',
            required: true
          },
          query: { 
            type: 'string', 
            description: 'Search query to filter contacts'
          },
          maxResults: { 
            type: 'number', 
            description: 'Maximum number of contacts to return (default: 10)'
          }
        },
        required: ['user_id']
      },
      handler: async ({ user_id, query, maxResults = 10 }) => {
        try {
          const accessToken = await getValidAccessToken(user_id);
          const graphClient = getGraphClient(accessToken);
          
          let endpoint = '/me/contacts';
          if (query) {
            endpoint = `/me/contacts?$filter=contains(displayName,'${query}') or contains(emailAddresses/any(e:e/address),'${query}')`;
          }
          
          const response = await graphClient
            .api(endpoint)
            .top(maxResults)
            .select('id,displayName,emailAddresses,businessPhones,mobilePhone')
            .orderBy('displayName')
            .get();
            
          return response.value;
        } catch (error) {
          logger.error('Error listing contacts:', error);
          throw new Error(`Failed to list contacts: ${error.message}`);
        }
      }
    }
  ];
  
  // Store user tokens
  // TODO: Use a persistent and secure store for tokens in production (e.g., Redis, encrypted DB)

  // Async error-handling wrapper
  function wrapAsync(fn) {
    return function(req, res, next) {
      Promise.resolve(fn(req, res, next)).catch(next);
    };
  }
  
  // Create the OAuth2 PKCE client
  const { AuthorizationCode, RefreshToken } = simpleOauth2;
  const oauth2 = new AuthorizationCode({
    client: {
      id: config.clientId
    },
    auth: {
      authorizeHost: authConfig.authorizeHost,
      authorizePath: authConfig.authorizePath,
      tokenHost: authConfig.tokenHost,
      tokenPath: authConfig.tokenPath
    },
    options: {
      authorizationMethod: 'body'
    }
  });
  
  // Initialize Microsoft Graph Client with access token
  function getGraphClient(accessToken) {
    return Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
  }
  
  // Authentication routes
  app.get('/auth/login', wrapAsync(async (req, res) => {
    // Generate and store a random state for CSRF protection
    const state = crypto.randomBytes(16).toString('hex');
    req.session.oauthState = state;
    // Create the authorization URL
    const authorizationUri = oauth2.authorizeURL({
      redirect_uri: config.redirectUri,
      scope: scopes.join(' '),
      state: state
    });
    // Redirect the user to the authorization URL
    res.redirect(authorizationUri);
  }));

  app.get('/auth/callback', wrapAsync(async (req, res) => {
    const { code, state } = req.query;
    // Verify state parameter for CSRF protection
    if (state !== req.session.oauthState) {
      return res.status(403).send('State validation failed');
    }
    // Exchange the authorization code for an access token
    const tokenParams = {
      code,
      redirect_uri: config.redirectUri,
      scope: scopes.join(' ')
    };
    const token = await oauth2.getToken(tokenParams);
    // Store the token in memory (in production, use a more secure storage method)
    const userId = crypto.randomBytes(16).toString('hex');
    userTokens[userId] = token;
    req.session.userId = userId;
    // Redirect to a success page
    res.redirect('/auth/success');
  }));

  app.get('/auth/success', (req, res) => {
    if (!req.session.userId || !userTokens[req.session.userId]) {
      return res.redirect('/auth/login');
    }
    
    res.send(`
      <html>
        <head>
          <title>Authentication Successful</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              margin-top: 50px;
            }
            .success {
              color: green;
              font-size: 24px;
              margin-bottom: 20px;
            }
          </style>
        </head>
        <body>
          <div class="success">✓ Successfully authenticated with Microsoft Outlook</div>
          <p>You can now close this window and use the MCP server.</p>
        </body>
      </html>
    `);
  });
  
  app.get('/auth/status', (req, res) => {
    const isAuthenticated = !!(req.session.userId && userTokens[req.session.userId]);
    
    res.json({
      authenticated: isAuthenticated,
      userId: isAuthenticated ? req.session.userId : null
    });
  });
  
  app.get('/auth/logout', (req, res) => {
    if (req.session.userId) {
      delete userTokens[req.session.userId];
      req.session.destroy();
    }
    
    res.redirect('/');
  });
  
  // Helper function to refresh the token if needed
  async function getValidAccessToken(userId) {
    if (!userTokens[userId]) {
      throw new Error('User not authenticated');
    }
    
    // Check if the token is about to expire (within 5 minutes)
    const tokenExpiresAt = userTokens[userId].token.expires_at;
    const now = new Date();
    const expiresIn = Math.floor((tokenExpiresAt - now) / 1000);
    
    // If token expires in less than 5 minutes, refresh it
    if (expiresIn < 300) {
      try {
        const refreshToken = userTokens[userId].token.refresh_token;
        
        if (!refreshToken) {
          throw new Error('No refresh token available');
        }
        
        // Create a RefreshToken client
        const refreshTokenClient = new RefreshToken({
          client: {
            id: config.clientId,
          },
          auth: {
            tokenHost: authConfig.tokenHost,
            tokenPath: authConfig.tokenPath
          }
        });
        
        // Refresh the token
        const newToken = await refreshTokenClient.refresh(refreshToken);
        userTokens[userId] = newToken;
        
        return newToken.token.access_token;
      } catch (error) {
        logger.error('Token refresh failed:', error);
        delete userTokens[userId];
        throw new Error('Authentication expired, please login again');
      }
    }
    
    return userTokens[userId].token.access_token;
  }
  
  // Initialize MCP Server
  const server = new McpServer({
    name: 'M365 MCP Server',
    version: '1.0.0'
  });
  
  // Make the MCP server instance available on the Express app for testing
  app.set('mcpServer', server);

  // Register authentication tool
  server.tool(
    'get_auth_status',
    z.object({
      user_id: z.string().optional().describe('User ID from authentication')
    }),
    async ({ user_id }) => {
      try {
        const isAuthenticated = !!(user_id && userTokens[user_id]);
        
        if (!isAuthenticated) {
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                authenticated: false,
                login_url: `http://localhost:${config.port}/auth/login`
              })
            }]
          };
        }
        
        // Check if we have a valid token
        try {
          await getValidAccessToken(user_id);
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({ authenticated: true })
            }]
          };
        } catch (error) {
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                authenticated: false,
                login_url: `http://localhost:${config.port}/auth/login`,
                error: 'Token expired or invalid'
              })
            }]
          };
        }
      } catch (error) {
        logger.error('Error getting auth status:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to check authentication status',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Register email tools
  server.tool(
    'list_emails',
    z.object({
      user_id: z.string().describe('User ID from authentication'),
      query: z.string().optional().describe('Search query to filter emails'),
      maxResults: z.number().min(1).max(50).default(10).describe('Maximum number of emails to return')
    }),
    async ({ user_id, query, maxResults }) => {
      try {
        const accessToken = await getValidAccessToken(user_id);
        const graphClient = getGraphClient(accessToken);
        
        let endpoint = '/me/messages';
        if (query) {
          endpoint = `/me/messages?$filter=contains(subject,'${query}')`;
        }
        
        const response = await graphClient
          .api(endpoint)
          .top(maxResults)
          .select('id,subject,bodyPreview,from,receivedDateTime')
          .orderBy('receivedDateTime DESC')
          .get();
          
        return {
          content: [{
            type: 'text',
            text: JSON.stringify(response.value)
          }]
        };
      } catch (error) {
        logger.error('Error listing emails:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to list emails',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Register email search tool
  server.tool(
    'search_emails',
    z.object({
      user_id: z.string().describe('User ID from authentication'),
      query: z.string().describe('Advanced search query (e.g., "from:example@gmail.com has:attachment")'),
      maxResults: z.number().min(1).max(50).default(10).describe('Maximum number of emails to return')
    }),
    async ({ user_id, query, maxResults }) => {
      try {
        const accessToken = await getValidAccessToken(user_id);
        const graphClient = getGraphClient(accessToken);
        
        const response = await graphClient
          .api('/me/messages')
          .search(query)
          .top(maxResults)
          .select('id,subject,bodyPreview,from,receivedDateTime,hasAttachments')
          .orderBy('receivedDateTime DESC')
          .get();
          
        return {
          content: [{
            type: 'text',
            text: JSON.stringify(response.value)
          }]
        };
      } catch (error) {
        logger.error('Error searching emails:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to search emails',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Register email retrieval tool
  server.tool(
    'get_email',
    z.object({
      user_id: z.string().describe('User ID from authentication'),
      id: z.string().describe('Email message ID'),
      format: z.enum(['html', 'text']).default('html').describe('Format of the email body')
    }),
    async ({ user_id, id, format }) => {
      try {
        const accessToken = await getValidAccessToken(user_id);
        const graphClient = getGraphClient(accessToken);
        
        const response = await graphClient
          .api(`/me/messages/${id}`)
          .select('id,subject,body,bodyPreview,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,hasAttachments')
          .get();
          
        const emailData = {
          ...response,
          body: format === 'html' ? response.body.content : response.bodyPreview
        };

        return {
          content: [{
            type: 'text',
            text: JSON.stringify(emailData)
          }]
        };
      } catch (error) {
        logger.error('Error getting email:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to get email',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Register email sending tool
  server.tool(
    'send_email',
    z.object({
      user_id: z.string().describe('User ID from authentication'),
      to: z.array(z.string().email()).describe('List of recipient email addresses'),
      subject: z.string().min(1).describe('Email subject'),
      body: z.string().min(1).describe('Email body content'),
      cc: z.array(z.string().email()).optional().describe('List of CC recipients'),
      bcc: z.array(z.string().email()).optional().describe('List of BCC recipients'),
      isHtml: z.boolean().default(true).describe('Whether the body is HTML')
    }),
    async ({ user_id, to, subject, body, cc, bcc, isHtml }) => {
      try {
        const accessToken = await getValidAccessToken(user_id);
        const graphClient = getGraphClient(accessToken);
        
        const message = {
          message: {
            subject,
            body: {
              contentType: isHtml ? 'HTML' : 'Text',
              content: body
            },
            toRecipients: to.map(email => ({
              emailAddress: { address: email }
            })),
            ccRecipients: cc?.map(email => ({
              emailAddress: { address: email }
            })) || [],
            bccRecipients: bcc?.map(email => ({
              emailAddress: { address: email }
            })) || []
          },
          saveToSentItems: true
        };
        
        await graphClient.api('/me/sendMail').post(message);
        
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: 'Email sent successfully'
            })
          }]
        };
      } catch (error) {
        logger.error('Error sending email:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to send email',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Register calendar tools
  server.tool(
    'list_events',
    z.object({
      user_id: z.string().describe('User ID from authentication'),
      timeMin: z.string().datetime().optional().describe('Start of time range (ISO 8601)'),
      timeMax: z.string().datetime().optional().describe('End of time range (ISO 8601)'),
      maxResults: z.number().min(1).max(50).default(10).describe('Maximum number of events')
    }),
    async ({ user_id, timeMin, timeMax, maxResults }) => {
      try {
        const accessToken = await getValidAccessToken(user_id);
        const graphClient = getGraphClient(accessToken);
        
        let endpoint = '/me/calendar/events';
        const params = [
          `$top=${maxResults}`,
          '$select=id,subject,start,end,location,organizer,attendees,webLink',
          '$orderby=start/dateTime'
        ];
        
        if (timeMin) params.push(`startDateTime=${encodeURIComponent(timeMin)}`);
        if (timeMax) params.push(`endDateTime=${encodeURIComponent(timeMax)}`);
        
        endpoint += '?' + params.join('&');
        
        const response = await graphClient.api(endpoint).get();
        
        return {
          content: [{
            type: 'text',
            text: JSON.stringify(response.value)
          }]
        };
      } catch (error) {
        logger.error('Error listing events:', error);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              error: 'Failed to list events',
              details: error.message
            })
          }],
          isError: true
        };
      }
    }
  );

  // Set up MCP HTTP server
  app.post('/v2/mcp',
    express.json(),
    wrapAsync(async (req, res) => {
      try {
        const result = await server.handleRequest(req.body);
        res.json(result);
      } catch (error) {
        res.status(500).json({
          error: 'Failed to process request',
          details: error.message
        });
      }
    })
  );

  return app;
}

// Only start the server if this file is run directly (ESM equivalent)
if (import.meta.url === `file://${process.argv[1]}`) {
  const app = createApp();
  const port = process.env.PORT || 3000;
  const serverInstance = app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });
  app.set('server', serverInstance);
  process.on('SIGTERM', () => {
    console.log('SIGTERM received. Shutting down gracefully');
    serverInstance.close(() => {
      console.log('Process terminated');
    });
  });
}

export { createApp };