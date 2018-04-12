import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import GraphCommand from '../../GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class GraphO365GroupRestoreCommand extends GraphCommand {

  public get name(): string {
    return commands.O365GROUP_RESTORE;
  }

  public get description(): string {
    return 'Restores a deleted Office 365 Group';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise | Promise<void> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Restoring Office 365 Group: ${args.options.id}...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/beta/directory/deleteditems/${args.options.id}/restore/`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            'accept': 'application/json;odata.metadata=none'
          }),
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
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
        description: 'The ID of the Office 365 Group to restore'
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
      `  ${chalk.yellow('Important:')} before using this command, connect to the Microsoft Graph,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:

    ${chalk.yellow('Attention:')} This command is based on a Microsoft Graph API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    To restore a deleted Office 365 Group, you have to first connect to
    the Microsoft Graph using the ${chalk.blue(commands.CONNECT)} command.

    If the specified ${chalk.grey('id')} doesn't refer to a deleted group, you will get
    a ${chalk.grey('File not found')} error.

  Examples:

    Restore group with ID ${chalk.grey('28beab62-7540-4db1-a23f-29a6018a3848')}.
      ${chalk.grey(config.delimiter)} ${this.name} --id 28beab62-7540-4db1-a23f-29a6018a3848
  `);
  }
}

module.exports = new GraphO365GroupRestoreCommand();