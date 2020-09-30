import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
    return `${commands.LIST_SET}`;
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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const body: any = {
      displayName: args.options.newName
    };

    const getListId = () => {
      if (args.options.name) {
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

    getListId().then(listId => {
      if (!listId) {
        return Promise.reject(`The list ${args.options.name} cannot be found`);
      }
      const requestOptions: any = {
        url: `${this.resource}/beta/me/todo/lists/${listId}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        body,
        json: true
      };

      return request.patch(requestOptions)
    }).then((): void => {
      if (this.verbose) {
        cmd.log(vorpal.chalk.green('DONE'));
      }

      cb();
    }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: `The ID of the list to update. Specify either id or name, not both`
      },
      {
        option: '-n, --name <name>',
        description: `The display name of the list to update. Specify either id or name, not both`
      },
      {
        option: '--newName <newName>',
        description: `The new name for the task list`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name && !args.options.id) {
        return 'Specify name or id of the list to update';
      }

      if (args.options.name && args.options.id) {
        return 'Specify either the name or the id of the list to update but not both'
      }

      if (!args.options.newName) {
        return 'Required option newName is missing'
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(`  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.
    
  Examples:

    Rename the list with ID ${chalk.grey("AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=")} to "My updated task list"
      m365 ${this.name} --id "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIhAAA=" --newName "My updated task list"

    Rename the list with name ${chalk.grey("My Task list")} to "My updated task list"
      m365 ${this.name} --name "My Task list" --newName "My updated task list"
`);
  }
}

module.exports = new TodoListSetCommand();