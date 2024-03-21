import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import ToDoCommand from '../../../base/ToDoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
  newName: string;
}

class TodoListSetCommand extends ToDoCommand {
  public get name(): string {
    return commands.LIST_SET;
  }

  public get description(): string {
    return 'Updates a Microsoft To Do task list';
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
        id: typeof args.options.id !== 'undefined',
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--newName <newName>'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['name', 'id'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0`;
    const data: any = {
      displayName: args.options.newName
    };

    try {
      const listId: string = await this.getListId(args);

      if (!listId) {
        throw `The list ${args.options.name} cannot be found`;
      }

      const requestOptions: any = {
        url: `${endpoint}/me/todo/lists/${listId}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        data,
        responseType: 'json'
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListId(args: CommandArgs): Promise<string> {
    const endpoint: string = `${this.resource}/v1.0`;
    if (args.options.id) {
      return args.options.id;
    }

    const requestOptions: any = {
      url: `${endpoint}/me/todo/lists?$filter=displayName eq '${escape(args.options.name as string)}'`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response.value && response.value.length === 1 ? response.value[0].id : null);
  }
}

export default new TodoListSetCommand();