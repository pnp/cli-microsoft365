import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  listName?: string;
  listId?: string;
  bodyContent?: string;
  bodyContentType?: string;
  dueDateTime?: string;
  importance?: string;
  reminderDateTime?: string;
  categories?: string;
  completedDateTime?: string;
  startDateTime?: string;
  status?: string;
}

class TodoTaskAddCommand extends GraphCommand {
  private static readonly allowedStatuses: string[] = ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'];

  public get name(): string {
    return commands.TASK_ADD;
  }

  public get description(): string {
    return 'Add a task to a Microsoft To Do task list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listName: typeof args.options.listName !== 'undefined',
        bodyContent: typeof args.options.bodyContent !== 'undefined',
        bodyContentType: args.options.bodyContentType,
        dueDateTime: typeof args.options.dueDateTime !== 'undefined',
        importance: args.options.importance,
        reminderDateTime: typeof args.options.reminderDateTime !== 'undefined',
        categories: typeof args.options.categories !== 'undefined',
        completedDateTime: typeof args.options.completedDateTime !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined',
        status: typeof args.options.status !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '--listName [listName]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--bodyContent [bodyContent]'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['text', 'html']
      },
      {
        option: '--dueDateTime [dueDateTime]'
      },
      {
        option: '--importance [importance]',
        autocomplete: ['low', 'normal', 'high']
      },
      {
        option: '--reminderDateTime [reminderDateTime]'
      },
      {
        option: '--categories [categories]'
      },
      {
        option: '--completedDateTime [completedDateTime]'
      },
      {
        option: '--startDateTime [startDateTime]'
      },
      {
        option: '--status [status]',
        autocomplete: TodoTaskAddCommand.allowedStatuses
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.bodyContentType && ['text', 'html'].indexOf(args.options.bodyContentType.toLowerCase()) === -1) {
          return `'${args.options.bodyContentType}' is not a valid value for the bodyContentType option. Allowed values are text|html`;
        }

        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance.toLowerCase()) === -1) {
          return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
        }

        if (args.options.dueDateTime && !validation.isValidISODateTime(args.options.dueDateTime)) {
          return `'${args.options.dueDateTime}' is not a valid ISO date string`;
        }

        if (args.options.reminderDateTime && !validation.isValidISODateTime(args.options.reminderDateTime)) {
          return `'${args.options.reminderDateTime}' is not a valid ISO date string`;
        }

        if (args.options.completedDateTime && !validation.isValidISODateTime(args.options.completedDateTime)) {
          return `'${args.options.completedDateTime}' is not a valid datetime.`;
        }

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return `'${args.options.startDateTime}' is not a valid datetime.`;
        }

        if (args.options.status && TodoTaskAddCommand.allowedStatuses.map(x => x.toLowerCase()).indexOf(args.options.status.toLowerCase()) === -1) {
          return `${args.options.status} is not a valid value for status. Valid values are ${TodoTaskAddCommand.allowedStatuses.join(', ')}`;
        }

        if (args.options.completedDateTime && args.options.status?.toLowerCase() !== 'completed') {
          return `The completedDateTime option can only be used when the status option is set to completed`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0`;

    try {
      const listId: string = await this.getTodoListId(args);

      const status = args.options.status && TodoTaskAddCommand.allowedStatuses.filter(x => x.toLowerCase() === args.options.status!.toLowerCase())[0];

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
      logger.log(res);
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

  private getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return Promise.resolve(args.options.listId);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.listName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<{ value: [{ id: string }] }>(requestOptions)
      .then(response => {
        const taskList: { id: string } | undefined = response.value[0];

        if (!taskList) {
          return Promise.reject(`The specified task list does not exist`);
        }

        return Promise.resolve(taskList.id);
      });
  }
}

module.exports = new TodoTaskAddCommand();