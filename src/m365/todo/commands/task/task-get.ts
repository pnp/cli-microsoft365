import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoTask } from '../../ToDoTask';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listName?: string;
  listId?: string;
}

class TodoTaskGetCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_GET;
  }

  public get description(): string {
    return 'Get a specific task from a Microsoft To Do task list';
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
        option: '-i, --id <id>'
      },
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

  private async getTodoListId(args: CommandArgs): Promise<string> {
    if (args.options.listId) {
      return args.options.listId;
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.listName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ id: string }] }>(requestOptions);

    const taskList = response.value[0];
    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listId: string = await this.getTodoListId(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/todo/lists/${listId}/tasks/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const item: ToDoTask = await request.get(requestOptions);

      if (args.options.output === 'json') {
        logger.log(item);
      }
      else {
        logger.log({
          id: item.id,
          title: item.title,
          status: item.status,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TodoTaskGetCommand();