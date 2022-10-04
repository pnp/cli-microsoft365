import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
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
  confirm?: boolean;
}

class TodoTaskRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TASK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft To Do task';
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
        listName: typeof args.options.listName !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        confirm: typeof args.options.confirm !== 'undefined'
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
      },
      {
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listName', 'listId']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const getToDoListId = async (): Promise<string | undefined> => {
      if (args.options.listName) {
        // Search list by its name
        const requestOptions: any = {
          url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.listName)}'`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          responseType: 'json'
        };
        const response: { value: { id: string }[] } = await request.get<{ value: { id: string }[] }>(requestOptions);
          
        return response.value && response.value.length === 1 ? response.value[0].id : undefined;  
      }

      return Promise.resolve(args.options.listId as string);
    };

    const removeToDoTask = async () => {
      try {
        const toDoListId: string | undefined = await getToDoListId();

        if (!toDoListId) {
          throw `The list ${args.options.listName} cannot be found`;
        }

        const requestOptions: any = {
          url: `${this.resource}/v1.0/me/todo/lists/${toDoListId}/tasks/${args.options.id}`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      } 
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeToDoTask();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the task ${args.options.id} from list ${args.options.listId || args.options.listName}?`
      });
      
      if (result.continue) {
        await removeToDoTask();
      }
    }
  }
}

module.exports = new TodoTaskRemoveCommand();