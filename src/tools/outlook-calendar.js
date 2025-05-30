import { zodToJsonSchema } from 'zod-to-json-schema';
import { z } from 'zod';
import { createGraphClient } from '../utils/graph-client.js';

export class OutlookCalendarTool {
  name = 'outlook-calendar';
  description = 'Tool for managing Outlook calendar events';
  inputSchema = zodToJsonSchema(z.object({
    action: z.enum(['list', 'create', 'update', 'delete']).describe('Action to perform'),
    eventId: z.string().optional().describe('Event ID for update/delete operations'),
    subject: z.string().optional().describe('Event subject'),
    body: z.string().optional().describe('Event body'),
    start: z.string().optional().describe('Event start time (ISO format)'),
    end: z.string().optional().describe('Event end time (ISO format)'),
    attendees: z.array(z.string()).optional().describe('List of attendee email addresses')
  }));
  #graphClient;

  constructor() {}

  initialize(config) {
    this.#graphClient = createGraphClient(config);
  }

  async call(args) {
    try {
      const { action, eventId, subject, body, start, end, attendees } = args;

      switch (action) {
        case 'list':
          const response = await this.#graphClient
            .api('/me/calendar/events')
            .select('id,subject,start,end,location,bodyPreview,webLink')
            .top(10)
            .orderby('start/dateTime')
            .get();
          return {
            content: [{
              type: 'text',
              text: JSON.stringify(response.value)
            }]
          };

        case 'create':
          const event = {
            subject,
            start: { dateTime: start, timeZone: 'UTC' },
            end: { dateTime: end, timeZone: 'UTC' },
            body: body ? { contentType: 'HTML', content: body } : undefined,
            attendees: attendees?.map(email => ({
              emailAddress: { address: email },
              type: 'required'
            }))
          };
          const createResponse = await this.#graphClient
            .api('/me/calendar/events')
            .post(event);
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                success: true,
                eventId: createResponse.id,
                event: createResponse
              })
            }]
          };

        case 'update':
          const updates = {};
          if (subject) updates.subject = subject;
          if (start) updates.start = { dateTime: start, timeZone: 'UTC' };
          if (end) updates.end = { dateTime: end, timeZone: 'UTC' };
          if (body) updates.body = { contentType: 'HTML', content: body };
          const updateResponse = await this.#graphClient
            .api(`/me/calendar/events/${eventId}`)
            .patch(updates);
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                success: true,
                event: updateResponse
              })
            }]
          };

        case 'delete':
          await this.#graphClient
            .api(`/me/calendar/events/${eventId}`)
            .delete();
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                success: true
              })
            }]
          };

        default:
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                error: 'Invalid action',
                details: 'Invalid action provided'
              })
            }],
            isError: true
          };
      }
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to perform action',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }
}
