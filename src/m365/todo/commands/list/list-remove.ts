import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  force?: boolean;
}

class TodoListRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.LIST_REMOVE;
  }

  public get description(): string {
    return 'Removes a Microsoft To Do task list';
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
        name: typeof args.options.name !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        force: typeof args.options.force !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['name', 'id'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeList(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the task list ${args.options.id || args.options.name}?` });

      if (result) {
        await this.removeList(args);
      }
    }
  }

  private async getListId(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id as string;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.name!)}'`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);

    return response.value && response.value.length === 1 ? response.value[0].id : null;
  }

  private async removeList(args: CommandArgs): Promise<void> {
    try {
      const listId: string = await this.getListId(args);

      if (!listId) {
        throw `The list ${args.options.name} cannot be found`;
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/me/todo/lists/${listId}`,
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

export default new TodoListRemoveCommand();