import { Logger } from '../../../../cli';
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
  newName: string;
}

class TodoListSetCommand extends GraphCommand {
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
    this.optionSets.push(['name', 'id']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const endpoint: string = `${this.resource}/v1.0`;
    const data: any = {
      displayName: args.options.newName
    };

    this
      .getListId(args)
      .then(listId => {
        if (!listId) {
          return Promise.reject(`The list ${args.options.name} cannot be found`);
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

        return request.patch(requestOptions);
      })
      .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getListId(args: CommandArgs): Promise<string> {
    const endpoint: string = `${this.resource}/v1.0`;
    if (args.options.id) {
      return Promise.resolve(args.options.id);
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

module.exports = new TodoListSetCommand();