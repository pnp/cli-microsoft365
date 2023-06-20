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
  id?: string;
  name?: string;
  confirm?: boolean;
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
        confirm: typeof args.options.confirm !== 'undefined'
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
        option: '--confirm'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['name', 'id'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeList(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the task list ${args.options.id || args.options.name}?`
      });

      if (result.continue) {
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

module.exports = new TodoListRemoveCommand();