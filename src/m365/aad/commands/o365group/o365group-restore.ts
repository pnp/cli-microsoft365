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
}

class AadO365GroupRestoreCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted Microsoft 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Restoring Microsoft 365 Group: ${args.options.id}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/directory/deleteditems/${args.options.id}/restore/`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
    };

    request
      .post(requestOptions)
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'The ID of the Microsoft 365 Group to restore'
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

    If the specified ${chalk.grey('id')} doesn't refer to a deleted group, you will get
    a ${chalk.grey('File not found')} error.

  Examples:

    Restore group with ID ${chalk.grey('28beab62-7540-4db1-a23f-29a6018a3848')}.
      ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848
  `);
  }
}

module.exports = new AadO365GroupRestoreCommand();