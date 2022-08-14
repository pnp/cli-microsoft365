import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listName?: string;
  listId?: string;
  title?: string;
  status?: string;
}

class TodoTaskSetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Update a task in a Microsoft To Do task list';
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
        status: typeof args.options.status !== 'undefined',
        title: typeof args.options.title !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-s, --status [status]',
        autocomplete: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']
      },
      {
        option: '--listName [listName]'
      },
      {
        option: '--listId [listId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.status &&
          args.options.status !== 'notStarted' &&
          args.options.status !== 'inProgress' &&
          args.options.status !== 'completed' &&
          args.options.status !== 'waitingOnOthers' &&
          args.options.status !== 'deferred') {
          return `${args.options.status} is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listName']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${this.resource}/v1.0`;
    const data = this.mapRequestBody(args.options);
    this
      .getTodoListId(args)
      .then((listId: string): Promise<any> => {
        const requestOptions: any = {
          url: `${endpoint}/me/todo/lists/${listId}/tasks/${encodeURIComponent(args.options.id)}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'Content-Type': 'application/json'
          },
          data: data,
          responseType: 'json'
        };

        return request.patch(requestOptions);
      })
      .then((res): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.status) {
      requestBody.status = options.status;
    }

    if (options.title) {
      requestBody.title = options.title;
    }

    return requestBody;
  }
}


module.exports = new TodoTaskSetCommand();