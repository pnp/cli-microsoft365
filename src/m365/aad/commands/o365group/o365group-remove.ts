import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class AadO365GroupRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_REMOVE;
  }

  public get description(): string {
    return 'Removes an Microsoft 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeGroup: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Removing Microsoft 365 Group: ${args.options.id}...`);
      }

      const requestOptions: any = {
        url: `${this.resource}/v1.0/groups/${args.options.id}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
    };

    if (args.options.confirm) {
      removeGroup();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the group ${args.options.id}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeGroup();
        }
      });
    }
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the Microsoft 365 Group to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the group'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required option id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    If the specified ${chalk.grey('id')} doesn't refer to an existing group, you will get
    a ${chalk.grey('Resource does not exist')} error.

  Examples:

    Remove group with ID ${chalk.grey('28beab62-7540-4db1-a23f-29a6018a3848')}. Will prompt
    for confirmation before removing the group
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848

    Remove group with ID ${chalk.grey('28beab62-7540-4db1-a23f-29a6018a3848')} without prompting
    for confirmation
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848 --confirm
  `);
  }
}

module.exports = new AadO365GroupRemoveCommand();