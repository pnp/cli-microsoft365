import { Cli, Logger } from '../../../../cli';
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
    this.#initValidators();
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

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.name && !args.options.id) {
          return 'Specify name or id of the list to remove';
        }

        if (args.options.name && args.options.id) {
          return 'Specify either the name or the id of the list to remove but not both';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const getListId = () => {
      if (args.options.name) {
        // Search list by its name
        const requestOptions: any = {
          url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.name)}'`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          responseType: 'json'
        };
        return request
          .get(requestOptions)
          .then((response: any) => response.value && response.value.length === 1 ? response.value[0].id : null);
      }

      return Promise.resolve(args.options.id);
    };

    const removeList = () => {
      getListId()
        .then(listId => {
          if (!listId) {
            return Promise.reject(`The list ${args.options.name} cannot be found`);
          }

          const requestOptions: any = {
            url: `${this.resource}/v1.0/me/todo/lists/${listId}`,
            headers: {
              accept: "application/json;odata.metadata=none"
            },
            responseType: 'json'
          };

          return request.delete(requestOptions);
        })
        .then(_ => cb(), (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeList();
    }
    else {
      Cli.prompt(
        {
          type: "confirm",
          name: "continue",
          default: false,
          message: `Are you sure you want to remove the task list ${args.options.id || args.options.name}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeList();
          }
        }
      );
    }
  }
}

module.exports = new TodoListRemoveCommand();