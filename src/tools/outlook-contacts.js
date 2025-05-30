import { Tool } from '@modelcontextprotocol/sdk/dist/cjs/server/tool.js';
import { z } from 'zod';
import { createGraphClient } from '../utils/graph-client.js';

export class OutlookContactsTool extends Tool {
  #graphClient;

  constructor() {
    super({
      name: 'outlook_contacts',
      description: 'Interact with Outlook contacts',
      version: '1.0.0'
    });

    // Register tool methods with their schemas
    this.registerMethod('list_contacts', {
      description: 'List contacts',
      parameters: z.object({
        maxResults: z.number().min(1).max(100).default(50).describe('Maximum number of contacts to return'),
        filter: z.string().optional().describe('OData filter query'),
        orderBy: z.string().default('displayName').describe('Sort order')
      }),
      handler: this.listContacts.bind(this)
    });

    this.registerMethod('create_contact', {
      description: 'Create a new contact',
      parameters: z.object({
        displayName: z.string().min(1).describe('Contact display name'),
        emailAddresses: z.array(z.string().email()).min(1).describe('List of email addresses'),
        businessPhones: z.array(z.string()).optional().default([]).describe('List of business phone numbers'),
        mobilePhone: z.string().optional().describe('Mobile phone number'),
        jobTitle: z.string().optional().describe('Job title'),
        companyName: z.string().optional().describe('Company name')
      }),
      handler: this.createContact.bind(this)
    });

    this.registerMethod('update_contact', {
      description: 'Update an existing contact',
      parameters: z.object({
        contactId: z.string().describe('ID of the contact to update'),
        displayName: z.string().optional().describe('New contact display name'),
        emailAddresses: z.array(z.string().email()).optional().describe('New list of email addresses'),
        businessPhones: z.array(z.string()).optional().describe('New list of business phone numbers'),
        mobilePhone: z.string().optional().describe('New mobile phone number'),
        jobTitle: z.string().optional().describe('New job title'),
        companyName: z.string().optional().describe('New company name')
      }),
      handler: this.updateContact.bind(this)
    });

    this.registerMethod('delete_contact', {
      description: 'Delete a contact',
      parameters: z.object({
        contactId: z.string().describe('ID of the contact to delete')
      }),
      handler: this.deleteContact.bind(this)
    });
  }

  initialize(config) {
    this.#graphClient = createGraphClient(config);
  }

  async listContacts(params) {
    try {
      const { maxResults = 50, filter, orderBy = 'displayName' } = params;

      let query = this.#graphClient
        .api('/me/contacts')
        .select('id,displayName,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName')
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
            error: 'Failed to list contacts',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async createContact(params) {
    try {
      const { displayName, emailAddresses, businessPhones = [], mobilePhone, jobTitle, companyName } = params;

      const contact = {
        displayName,
        emailAddresses: emailAddresses.map(email => ({
          address: email,
          name: displayName
        })),
        businessPhones,
        mobilePhone,
        jobTitle,
        companyName
      };

      const response = await this.#graphClient
        .api('/me/contacts')
        .post(contact);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            contactId: response.id,
            contact: response
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to create contact',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async updateContact(params) {
    try {
      const { contactId, ...updates } = params;

      const response = await this.#graphClient
        .api(`/me/contacts/${contactId}`)
        .patch(updates);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            contact: response
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to update contact',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async deleteContact(params) {
    try {
      const { contactId } = params;

      await this.#graphClient
        .api(`/me/contacts/${contactId}`)
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
            error: 'Failed to delete contact',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }
}
