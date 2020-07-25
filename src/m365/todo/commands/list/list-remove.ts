import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
    return `${commands.LIST_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a Microsoft To Do task list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const getListId = () => {
      if (args.options.name) {
        // Search list by its name
        const requestOptions: any = {
          url: `${this.resource}/beta/me/todo/lists?$filter=displayName eq '${escape(args.options.name)}'`,
          headers: {
            accept: "application/json;odata.metadata=none"
          },
          json: true
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
            url: `${this.resource}/beta/me/todo/lists/${listId}`,
            headers: {
              accept: "application/json;odata.metadata=none"
            },
            json: true
          };

          return request.delete(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeList();
    }
    else {
      cmd.prompt(
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name [name]',
        description: `The name of the task list to remove. Specify either id or name but not both`
      },
      {
        option: '-i, --id [id]',
        description: `The ID of the task list to remove. Specify either id or name but not both`
      },
      {
        option: '--confirm',
        description: `Don't prompt for confirming removing the task list`
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name && !args.options.id) {
        return 'Specify name or id of the list to remove';
      }

      if (args.options.name && args.options.id) {
        return 'Specify either the name or the id of the list to remove but not both'
      }

      return true;
    };
  }
}

module.exports = new TodoListRemoveCommand();