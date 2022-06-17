import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoTask } from '../../ToDoTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  listName?: string;
  listId?: string;
}

class TodoTaskListCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_LIST;
  }

  public get description(): string {
    return 'List tasks from a Microsoft To Do task list';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'status', 'createdDateTime', 'lastModifiedDateTime'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listName: typeof args.options.listName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        if (args.options.listId && args.options.listName) {
          return 'Specify listId or listName, but not both';
        }

        if (!args.options.listId && !args.options.listName) {
          return 'Specify listId or listName';
        }

        return true;
      }
    );
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTodoListId(args)
      .then((listId: string): Promise<any> => {
        const endpoint: string = `${this.resource}/v1.0/me/todo/lists/${listId}/tasks`;
        return odata.getAllItems(endpoint);
      })
      .then((items: ToDoTask[]): void => {
        if (args.options.output === 'json') {
          logger.log(items);
        }
        else {
          logger.log(items.map(m => {
            return {
              id: m.id,
              title: m.title,
              status: m.status,
              createdDateTime: m.createdDateTime,
              lastModifiedDateTime: m.lastModifiedDateTime
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TodoTaskListCommand();