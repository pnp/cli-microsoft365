import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listName = typeof args.options.listName !== 'undefined';
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const getToDoListId = (): Promise<string | undefined> => {
      if (args.options.listName) {
        // Search list by its name
        const requestOptions: any = {
          url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${escape(args.options.listName)}'`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          responseType: 'json'
        };
        return request
          .get<{ value: { id: string }[] }>(requestOptions)
          .then((response): string | undefined =>
            response.value && response.value.length === 1 ? response.value[0].id : undefined);
      }

      return Promise.resolve(args.options.listId as string);
    };

    const removeToDoTask = () => {
      getToDoListId()
        .then(toDoListId => {
          if (!toDoListId) {
            return Promise.reject(`The list ${args.options.listName} cannot be found`);
          }

          const requestOptions: any = {
            url: `${this.resource}/v1.0/me/todo/lists/${toDoListId}/tasks/${args.options.id}`,
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
      removeToDoTask();
    }
    else {
      Cli.prompt(
        {
          type: "confirm",
          name: "continue",
          default: false,
          message: `Are you sure you want to remove the task ${args.options.id} from list ${args.options.listId || args.options.listName}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeToDoTask();
          }
        }
      );
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.listName && !args.options.listId) {
      return 'Specify name or id of the To Do list';
    }

    if (args.options.listName && args.options.listId) {
      return 'Specify either the name or the id of the To Do list but not both';
    }

    return true;
  }
}

module.exports = new TodoTaskRemoveCommand();