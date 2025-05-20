# MCP API Documentation

This document provides detailed information about the Model Context Protocol (MCP) API endpoints available in the Outlook MCP Server.

## Base URL

All API requests should be sent to:

```
http://localhost:3000/v2/mcp  (or your custom host/port)
```

## Authentication Flow

Before making API calls, you need to authenticate with Microsoft. The flow is as follows:

1. Check authentication status using `get_auth_status`
2. If not authenticated, direct user to the login URL
3. After login, get user_id for subsequent requests
4. Include user_id in all API requests

## API Tools

### Authentication Tools

#### `get_auth_status`

Check if the user is authenticated with Microsoft Outlook.

**Parameters:**
- None

**Response:**
```json
{
  "authenticated": true,
  "user_id": "abc123def456"
}
```

Or if not authenticated:
```json
{
  "authenticated": false,
  "login_url": "http://localhost:3000/auth/login"
}
```

### Email Tools

#### `list_emails`

List recent emails from the Outlook inbox.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `query` (string, optional): Search query to filter emails
- `maxResults` (number, optional): Maximum number of emails to return (default: 10)

**Response:**
```json
[
  {
    "id": "AAMkADRm...",
    "subject": "Meeting tomorrow",
    "bodyPreview": "Let's meet to discuss...",
    "from": {
      "emailAddress": {
        "name": "John Doe",
        "address": "john@example.com"
      }
    },
    "receivedDateTime": "2025-05-19T15:30:00Z"
  },
  ...
]
```

#### `search_emails`

Search emails with advanced query.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `query` (string, required): Outlook search query (e.g., "from:example@gmail.com has:attachment")
- `maxResults` (number, optional): Maximum number of emails to return (default: 10)

**Response:**
Same format as `list_emails`.

#### `get_email`

Get a specific email by ID.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `id` (string, required): Email ID
- `format` (string, optional): Format to retrieve the email (html or text, default: html)

**Response:**
```json
{
  "id": "AAMkADRm...",
  "subject": "Meeting tomorrow",
  "body": {
    "contentType": "html",
    "content": "<html><body>Let's meet to discuss...</body></html>"
  },
  "from": {
    "emailAddress": {
      "name": "John Doe",
      "address": "john@example.com"
    }
  },
  "toRecipients": [
    {
      "emailAddress": {
        "name": "Jane Smith",
        "address": "jane@example.com"
      }
    }
  ],
  "receivedDateTime": "2025-05-19T15:30:00Z",
  "hasAttachments": false
}
```

#### `send_email`

Send a new email.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `to` (string, required): Recipient email address
- `subject` (string, required): Email subject
- `body` (string, required): Email body (can include HTML)
- `cc` (string, optional): CC recipients (comma-separated)
- `bcc` (string, optional): BCC recipients (comma-separated)
- `isHtml` (boolean, optional): Whether the body is HTML (default: true)

**Response:**
```json
{
  "success": true,
  "message": "Email sent successfully"
}
```

### Calendar Tools

#### `list_events`

List upcoming calendar events.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `timeMin` (string, optional): Start time in ISO format (default: now)
- `timeMax` (string, optional): End time in ISO format
- `maxResults` (number, optional): Maximum number of events to return (default: 10)

**Response:**
```json
[
  {
    "id": "AAMkADRm...",
    "subject": "Weekly team meeting",
    "bodyPreview": "Agenda for this week...",
    "start": {
      "dateTime": "2025-05-20T09:00:00Z",
      "timeZone": "UTC"
    },
    "end": {
      "dateTime": "2025-05-20T10:00:00Z",
      "timeZone": "UTC"
    },
    "location": {
      "displayName": "Conference Room 3"
    },
    "organizer": {
      "emailAddress": {
        "name": "John Doe",
        "address": "john@example.com"
      }
    },
    "attendees": [
      {
        "type": "required",
        "emailAddress": {
          "name": "Jane Smith",
          "address": "jane@example.com"
        }
      }
    ]
  },
  ...
]
```

#### `create_event`

Create a new calendar event.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `subject` (string, required): Event title
- `start` (string, required): Start time in ISO format
- `end` (string, required): End time in ISO format
- `location` (string, optional): Event location
- `description` (string, optional): Event description
- `attendees` (array, optional): List of attendee email addresses
- `isOnlineMeeting` (boolean, optional): Whether this is an online meeting (default: false)

**Response:**
```json
{
  "id": "AAMkADRm...",
  "subject": "Project Review",
  "start": {
    "dateTime": "2025-05-22T14:00:00Z",
    "timeZone": "UTC"
  },
  "end": {
    "dateTime": "2025-05-22T15:00:00Z",
    "timeZone": "UTC"
  },
  "location": {
    "displayName": "Conference Room 2"
  },
  "body": {
    "contentType": "html",
    "content": "<html><body>Review project progress</body></html>"
  },
  "attendees": [
    {
      "type": "required",
      "emailAddress": {
        "name": "Jane Smith",
        "address": "jane@example.com"
      }
    }
  ],
  "isOnlineMeeting": true,
  "onlineMeeting": {
    "joinUrl": "https://teams.microsoft.com/l/meetup-join/..."
  }
}
```

### Contacts Tools

#### `list_contacts`

List contacts from Outlook.

**Parameters:**
- `user_id` (string, required): User ID from authentication
- `query` (string, optional): Search query to filter contacts
- `maxResults` (number, optional): Maximum number of contacts to return (default: 10)

**Response:**
```json
[
  {
    "id": "AAMkADRm...",
    "displayName": "John Doe",
    "emailAddresses": [
      {
        "name": "John Doe",
        "address": "john@example.com"
      }
    ],
    "businessPhones": [
      "+1 (555) 123-4567"
    ],
    "mobilePhone": "+1 (555) 987-6543"
  },
  ...
]
```

## Error Handling

The API returns standard HTTP status codes and error messages in the following format:

```json
{
  "error": {
    "message": "Error message",
    "code": "ERROR_CODE"
  }
}
```

Common error codes:
- `AUTHENTICATION_REQUIRED`: User is not authenticated
- `INVALID_PARAMETERS`: Missing or invalid parameters
- `PERMISSION_DENIED`: Insufficient permissions
- `RESOURCE_NOT_FOUND`: Requested resource not found
- `INTERNAL_SERVER_ERROR`: Server error

## Rate Limiting

The API currently relies on Microsoft Graph API's rate limits. Please refer to the [Microsoft Graph API documentation](https://learn.microsoft.com/en-us/graph/throttling) for details.