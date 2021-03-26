import { Logger } from '../../../../cli';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
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
          url: `${this.resource}/beta/me/todo/lists/${listId}`,
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
    if (args.options.id) {
      return Promise.resolve(args.options.id);
    }

    const requestOptions: any = {
      url: `${this.resource}/beta/me/todo/lists?$filter=displayName eq '${escape(args.options.name as string)}'`,
      headers: {
        accept: "application/json;odata.metadata=none"
      },
      responseType: 'json'
    };

    return request
      .get(requestOptions)
      .then((response: any) => response.value && response.value.length === 1 ? response.value[0].id : null);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '--newName <newName>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.name && !args.options.id) {
      return 'Specify name or id of the list to update';
    }

    if (args.options.name && args.options.id) {
      return 'Specify either the name or the id of the list to update but not both';
    }

    if (!args.options.newName) {
      return 'Required option newName is missing';
    }

    return true;
  }
}

module.exports = new TodoListSetCommand();