import { Logger } from '../../../../cli/Logger.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphDelegatedCommand from '../../../base/GraphDelegatedCommand.js';
import commands from '../../commands.js';

const allowedStatuses = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  title: z.string().optional().alias('t'),
  status: z.string().refine(val => allowedStatuses.includes(val as typeof allowedStatuses[number]), {
    message: `The value is not valid for status. Allowed values are ${allowedStatuses.join('|')}`
  }).optional().alias('s'),
  listName: z.string().optional(),
  listId: z.string().optional(),
  bodyContent: z.string().optional(),
  bodyContentType: z.string().refine(val => ['text', 'html'].includes(val.toLowerCase()), {
    message: 'The value is not valid for bodyContentType. Allowed values are text|html.'
  }).optional(),
  dueDateTime: z.string().refine(val => validation.isValidISODateTime(val), {
    message: 'The value is not a valid ISO date string.'
  }).optional(),
  importance: z.string().refine(val => ['low', 'normal', 'high'].includes(val.toLowerCase()), {
    message: 'The value is not valid for importance. Allowed values are low|normal|high.'
  }).optional(),
  reminderDateTime: z.string().refine(val => validation.isValidISODateTime(val), {
    message: 'The value is not a valid ISO date string.'
  }).optional(),
  categories: z.string().optional(),
  completedDateTime: z.string().refine(val => validation.isValidISODateTime(val), {
    message: 'The value is not a valid datetime.'
  }).optional(),
  startDateTime: z.string().refine(val => validation.isValidISODateTime(val), {
    message: 'The value is not a valid datetime.'
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoTaskSetCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Updates a task in a Microsoft To Do task list';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema.refine(opts => [opts.listId, opts.listName].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'listId' or 'listName', but not both.`,
      params: { customCode: 'optionSet', options: ['listId', 'listName'] }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0`;
    const data = this.mapRequestBody(args.options);

    try {
      const listId: string = await this.getTodoListId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/me/todo/lists/${listId}/tasks/${formatting.encodeQueryParameter(args.options.id)}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json'
        },
        data: data,
        responseType: 'json'
      };

      const res = await request.patch<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTodoListId(options: Options): Promise<string> {
    if (options.listId) {
      return options.listId;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.listName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ id: string }] }>(requestOptions);
    const taskList: { id: string } | undefined = response.value[0];

    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }

  private getDateTimeTimeZone(dateTime: string): { dateTime: string, timeZone: string } {
    return {
      dateTime: dateTime,
      timeZone: 'Etc/GMT'
    };
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.status) {
      requestBody.status = options.status;
    }

    if (options.title) {
      requestBody.title = options.title;
    }

    if (options.importance) {
      requestBody.importance = options.importance.toLowerCase();
    }

    if (options.bodyContentType || options.bodyContent) {
      requestBody.body = {
        content: options.bodyContent,
        contentType: options.bodyContentType?.toLowerCase() || 'text'
      };
    }

    if (options.dueDateTime) {
      requestBody.dueDateTime = this.getDateTimeTimeZone(options.dueDateTime);
    }

    if (options.reminderDateTime) {
      requestBody.reminderDateTime = this.getDateTimeTimeZone(options.reminderDateTime);
    }

    if (options.categories) {
      requestBody.categories = options.categories.split(',');
    }

    if (options.completedDateTime) {
      requestBody.completedDateTime = this.getDateTimeTimeZone(options.completedDateTime);
    }

    if (options.startDateTime) {
      requestBody.startDateTime = this.getDateTimeTimeZone(options.startDateTime);
    }

    return requestBody;
  }
}

export default new TodoTaskSetCommand();