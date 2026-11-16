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
  title: z.string().alias('t'),
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
  }).optional(),
  status: z.string().refine(val => allowedStatuses.some(status => status.toLowerCase() === val.toLowerCase()), {
    message: `The value is not valid for status. Valid values are ${allowedStatuses.join(', ')}`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TodoTaskAddCommand extends GraphDelegatedCommand {
  public get name(): string {
    return commands.TASK_ADD;
  }

  public get description(): string {
    return 'Adds a task to a Microsoft To Do list';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema.refine(opts => [opts.listId, opts.listName].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'listId' or 'listName', but not both.`,
      params: { customCode: 'optionSet', options: ['listId', 'listName'] }
    }).refine(opts => !opts.completedDateTime || opts.status?.toLowerCase() === 'completed', {
      message: 'The completedDateTime option can only be used when the status option is set to completed.',
      path: ['completedDateTime'],
      params: { customCode: 'required' }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0`;

    try {
      const listId: string = await this.getTodoListId(args);

      const status = args.options.status && allowedStatuses.find(x => x.toLowerCase() === args.options.status!.toLowerCase());

      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/me/todo/lists/${listId}/tasks`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json'
        },
        data: {
          title: args.options.title,
          body: {
            content: args.options.bodyContent,
            contentType: args.options.bodyContentType?.toLowerCase() || 'text'
          },
          importance: args.options.importance?.toLowerCase(),
          dueDateTime: this.getDateTimeTimeZone(args.options.dueDateTime),
          reminderDateTime: this.getDateTimeTimeZone(args.options.reminderDateTime),
          categories: args.options.categories?.split(','),
          completedDateTime: this.getDateTimeTimeZone(args.options.completedDateTime),
          startDateTime: this.getDateTimeTimeZone(args.options.startDateTime),
          status: status
        },
        responseType: 'json'
      };

      const res = await request.post<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getDateTimeTimeZone(dateTime: string | undefined): { dateTime: string, timeZone: string } | undefined {
    if (!dateTime) {
      return undefined;
    }

    return {
      dateTime: dateTime,
      timeZone: 'Etc/GMT'
    };
  }

  private async getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return args.options.listId;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.listName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: any = await request.get<{ value: [{ id: string }] }>(requestOptions);
    const taskList: { id: string } | undefined = response.value[0];

    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }
}

export default new TodoTaskAddCommand();