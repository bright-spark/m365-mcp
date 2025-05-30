import { Tool } from '@modelcontextprotocol/sdk/dist/cjs/server/tool.js';
import { z } from 'zod';
import { createGraphClient } from '../utils/graph-client.js';

export class OutlookEmailTool extends Tool {
  #graphClient;

  constructor() {
    super({
      name: 'outlook_email',
      description: 'Interact with Outlook email',
      version: '1.0.0'
    });

    // Register tool methods with their schemas
    this.registerMethod('list_emails', {
      description: 'List emails from a specified folder',
      parameters: z.object({
        folder: z.string().default('inbox').describe('Mail folder name'),
        maxResults: z.number().min(1).max(50).default(10).describe('Maximum number of emails to return'),
        filter: z.string().optional().describe('OData filter query'),
        orderBy: z.string().default('receivedDateTime desc').describe('Sort order')
      }),
      handler: this.listEmails.bind(this)
    });

    this.registerMethod('send_email', {
      description: 'Send a new email',
      parameters: z.object({
        subject: z.string().min(1).describe('Email subject'),
        body: z.string().min(1).describe('Email body in HTML format'),
        toRecipients: z.array(z.string().email()).min(1).describe('List of recipient email addresses'),
        ccRecipients: z.array(z.string().email()).optional().default([]).describe('List of CC recipient email addresses'),
        bccRecipients: z.array(z.string().email()).optional().default([]).describe('List of BCC recipient email addresses'),
        attachments: z.array(z.object({
          name: z.string().describe('Attachment filename'),
          content: z.string().describe('Base64 encoded file content')
        })).optional().default([]).describe('List of file attachments')
      }),
      handler: this.sendEmail.bind(this)
    });

    this.registerMethod('get_email', {
      description: 'Get details of a specific email',
      parameters: z.object({
        messageId: z.string().describe('ID of the email message')
      }),
      handler: this.getEmail.bind(this)
    });

    this.registerMethod('move_email', {
      description: 'Move an email to a different folder',
      parameters: z.object({
        messageId: z.string().describe('ID of the email message'),
        destinationFolder: z.string().describe('ID of the destination folder')
      }),
      handler: this.moveEmail.bind(this)
    });

    this.registerMethod('delete_email', {
      description: 'Delete an email',
      parameters: z.object({
        messageId: z.string().describe('ID of the email message')
      }),
      handler: this.deleteEmail.bind(this)
    });

    this.registerMethod('list_folders', {
      description: 'List mail folders',
      parameters: z.object({
        maxResults: z.number().min(1).max(100).default(50).describe('Maximum number of folders to return')
      }),
      handler: this.listFolders.bind(this)
    });
  }

  initialize(config) {
    this.#graphClient = createGraphClient(config);
  }

  async listEmails(params) {
    try {
      const { folder = 'inbox', maxResults = 10, filter, orderBy = 'receivedDateTime desc' } = params;

      let query = this.#graphClient
        .api(`/me/mailFolders/${folder}/messages`)
        .select('id,subject,sender,receivedDateTime,bodyPreview,hasAttachments,webLink')
        .top(maxResults)
        .orderby(orderBy);

      if (filter) {
        query = query.filter(filter);
      }

      const response = await query.get();
      return {
        content: [{
          type: 'text',
          text: JSON.stringify(response.value)
        }]
      };
    } catch (error) {
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

  async sendEmail(params) {
    try {
      const { subject, body, toRecipients, ccRecipients = [], bccRecipients = [], attachments = [] } = params;

      const message = {
        subject,
        body: {
          contentType: 'HTML',
          content: body
        },
        toRecipients: toRecipients.map(email => ({
          emailAddress: { address: email }
        })),
        ccRecipients: ccRecipients.map(email => ({
          emailAddress: { address: email }
        })),
        bccRecipients: bccRecipients.map(email => ({
          emailAddress: { address: email }
        })),
        attachments: attachments.map(attachment => ({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: attachment.name,
          contentBytes: attachment.content
        }))
      };

      const response = await this.#graphClient
        .api('/me/sendMail')
        .post({ message });

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.id
          })
        }]
      };
    } catch (error) {
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

  async getEmail(params) {
    try {
      const { messageId } = params;

      const response = await this.#graphClient
        .api(`/me/messages/${messageId}`)
        .select('id,subject,sender,receivedDateTime,body,toRecipients,ccRecipients,bccRecipients,hasAttachments,webLink')
        .get();

      return {
        content: [{
          type: 'text',
          text: JSON.stringify(response)
        }]
      };
    } catch (error) {
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

  async moveEmail(params) {
    try {
      const { messageId, destinationFolder } = params;

      const response = await this.#graphClient
        .api(`/me/messages/${messageId}/move`)
        .post({
          destinationId: destinationFolder
        });

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            newMessageId: response.id
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to move email',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async deleteEmail(params) {
    try {
      const { messageId } = params;

      await this.#graphClient
        .api(`/me/messages/${messageId}`)
        .delete();

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to delete email',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async listFolders(params) {
    try {
      const { maxResults = 50 } = params;

      const response = await this.#graphClient
        .api('/me/mailFolders')
        .select('id,displayName,parentFolderId,childFolderCount,totalItemCount')
        .top(maxResults)
        .get();

      return {
        content: [{
          type: 'text',
          text: JSON.stringify(response.value)
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to list folders',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }
}
