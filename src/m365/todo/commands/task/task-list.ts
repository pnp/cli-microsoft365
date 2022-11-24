import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { odata } from '../../../../utils/odata';
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
    this.#initOptionSets();
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

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listName'] });
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listId: string = await this.getTodoListId(args);
      const endpoint: string = `${this.resource}/v1.0/me/todo/lists/${listId}/tasks`;
      const items: ToDoTask[] = await odata.getAllItems(endpoint);

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
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TodoTaskListCommand();