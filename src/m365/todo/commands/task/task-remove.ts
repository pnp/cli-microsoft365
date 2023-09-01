import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listName?: string;
  listId?: string;
  force?: boolean;
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
        force: typeof args.options.force !== 'undefined'
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
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listName', 'listId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeToDoTask(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the task ${args.options.id} from list ${args.options.listId || args.options.listName}?`
      });

      if (result.continue) {
        await this.removeToDoTask(args);
      }
    }
  }

  private async getToDoListId(options: GlobalOptions): Promise<string | undefined> {
    if (options.listName) {
      // Search list by its name
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(options.listName)}'`,
        headers: {
          accept: "application/json;odata.metadata=none"
        },
        responseType: 'json'
      };
      const response: { value: { id: string }[] } = await request.get<{ value: { id: string }[] }>(requestOptions);

      return response.value && response.value.length === 1 ? response.value[0].id : undefined;
    }

    return options.listId as string;
  }

  private async removeToDoTask(args: CommandArgs): Promise<void> {
    try {
      const toDoListId: string | undefined = await this.getToDoListId(args.options);

      if (!toDoListId) {
        throw `The list ${args.options.listName} cannot be found`;
      }

      const requestOptions: CliRequestOptions = {
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
  }
}

export default new TodoTaskRemoveCommand();