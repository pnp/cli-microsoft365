import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { ToDoList } from '../../ToDoList';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  name?: string;
}


class TodoListGetCommand extends GraphCommand {
  public get name(): string {
    return commands.LIST_GET;
  }


  public get description(): string {
    return 'Gets a specific list of Microsoft To Do task lists';
  }

  public defaultProperties(): string[] | undefined {
    return ['displayName', 'id'];
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
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const item = await this.getList(args.options);
      logger.log(item);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getList(options: Options): Promise<any> {
    const requestOptions: AxiosRequestConfig = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.id) {
      requestOptions.url = `${this.resource}/v1.0/me/todo/lists/${options.id}`;
      const result = await request.get<ToDoList>(requestOptions);
      return result;
    }

    requestOptions.url = `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(options.name as string)}'`;
    const result = await request.get<{ "value": ToDoList[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified list '${options.name}' does not exist.`;
    }

    if (result.value.length > 1) {
      throw `Multiple lists with name '${options.name}' found: ${result.value.map(x => x.id).join(',')}`;
    }

    return result.value[0];
  }
}

module.exports = new TodoListGetCommand();