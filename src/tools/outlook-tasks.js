import { Tool } from '@modelcontextprotocol/sdk/dist/cjs/server/tool.js';
import { z } from 'zod';
import { createGraphClient } from '../utils/graph-client.js';

export class OutlookTasksTool extends Tool {
  #graphClient;

  constructor() {
    super({
      name: 'outlook_tasks',
      description: 'Interact with Outlook tasks',
      version: '1.0.0'
    });

    // Register tool methods with their schemas
    this.registerMethod('list_tasks', {
      description: 'List tasks from a specified list',
      parameters: z.object({
        listId: z.string().default('default').describe('ID of the task list'),
        maxResults: z.number().min(1).max(100).default(50).describe('Maximum number of tasks to return'),
        filter: z.string().optional().describe('OData filter query'),
        orderBy: z.string().default('createdDateTime desc').describe('Sort order')
      }),
      handler: this.listTasks.bind(this)
    });

    this.registerMethod('create_task', {
      description: 'Create a new task',
      parameters: z.object({
        listId: z.string().default('default').describe('ID of the task list'),
        title: z.string().min(1).describe('Task title'),
        dueDateTime: z.string().datetime().optional().describe('Due date and time in ISO format'),
        importance: z.enum(['low', 'normal', 'high']).default('normal').describe('Task importance'),
        body: z.string().optional().describe('Task description')
      }),
      handler: this.createTask.bind(this)
    });

    this.registerMethod('update_task', {
      description: 'Update an existing task',
      parameters: z.object({
        listId: z.string().default('default').describe('ID of the task list'),
        taskId: z.string().describe('ID of the task to update'),
        title: z.string().optional().describe('New task title'),
        status: z.enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']).optional().describe('Task status'),
        importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
        dueDateTime: z.string().datetime().optional().describe('New due date and time in ISO format'),
        body: z.string().optional().describe('New task description')
      }),
      handler: this.updateTask.bind(this)
    });

    this.registerMethod('delete_task', {
      description: 'Delete a task',
      parameters: z.object({
        listId: z.string().default('default').describe('ID of the task list'),
        taskId: z.string().describe('ID of the task to delete')
      }),
      handler: this.deleteTask.bind(this)
    });

    this.registerMethod('list_task_lists', {
      description: 'List task lists',
      parameters: z.object({
        maxResults: z.number().min(1).max(100).default(50).describe('Maximum number of task lists to return')
      }),
      handler: this.listTaskLists.bind(this)
    });
  }

  initialize(config) {
    this.#graphClient = createGraphClient(config);
  }

  async listTasks(params) {
    try {
      const { listId = 'default', maxResults = 50, filter, orderBy = 'createdDateTime desc' } = params;

      let query = this.#graphClient
        .api(`/me/todo/lists/${listId}/tasks`)
        .select('id,title,status,importance,dueDateTime,bodyLastModifiedDateTime,body')
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
            error: 'Failed to list tasks',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async createTask(params) {
    try {
      const { listId = 'default', title, dueDateTime, importance = 'normal', body } = params;

      const task = {
        title,
        dueDateTime: dueDateTime ? {
          dateTime: dueDateTime,
          timeZone: 'UTC'
        } : undefined,
        importance,
        body: body ? {
          content: body,
          contentType: 'text'
        } : undefined
      };

      const response = await this.#graphClient
        .api(`/me/todo/lists/${listId}/tasks`)
        .post(task);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            taskId: response.id,
            task: response
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to create task',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async updateTask(params) {
    try {
      const { listId = 'default', taskId, ...updates } = params;

      const task = {};
      if (updates.title) task.title = updates.title;
      if (updates.status) task.status = updates.status;
      if (updates.importance) task.importance = updates.importance;
      if (updates.dueDateTime) {
        task.dueDateTime = {
          dateTime: updates.dueDateTime,
          timeZone: 'UTC'
        };
      }
      if (updates.body) {
        task.body = {
          content: updates.body,
          contentType: 'text'
        };
      }

      const response = await this.#graphClient
        .api(`/me/todo/lists/${listId}/tasks/${taskId}`)
        .patch(task);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            task: response
          })
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            error: 'Failed to update task',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async deleteTask(params) {
    try {
      const { listId = 'default', taskId } = params;

      await this.#graphClient
        .api(`/me/todo/lists/${listId}/tasks/${taskId}`)
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
            error: 'Failed to delete task',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }

  async listTaskLists(params) {
    try {
      const { maxResults = 50 } = params;

      const response = await this.#graphClient
        .api('/me/todo/lists')
        .select('id,displayName,isOwner,isShared')
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
            error: 'Failed to list task lists',
            details: error.message
          })
        }],
        isError: true
      };
    }
  }
}
