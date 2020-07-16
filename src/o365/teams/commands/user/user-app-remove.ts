import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  userId: string;
  confirm?: boolean;
}

class TeamsUserAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_USER_APP_REMOVE}`;
  }

  public get description(): string {
    return 'Uninstall an app from the personal scope of the specified user.';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!!args.options.confirm).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeApp: () => void = (): void => {
      const endpoint: string = `${this.resource}/beta`

      const requestOptions: any = {
        url: `${endpoint}/users/${args.options.userId}/teamwork/installedApps/${args.options.appId}`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        json: true
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
    }

    if (args.options.confirm) {
      removeApp();
    }
    else {
      cmd.prompt(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the app with id ${args.options.appId} for user ${args.options.userId}?`
        },
        (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeApp();
          }
        }
      );
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--appId <appId>',
        description: 'The unique id of the app instance installed for the user'
      },
      {
        option: '--userId <userId>',
        description: 'The ID of the user to uninstall the app for'
      },
      {
        option: '--confirm',
        description: 'Confirm removal of app for user'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId) {
        return 'Required parameter appId missing';
      }

      if (!args.options.userId) {
        return 'Required parameter userId missing';
      }

      if (!Utils.isValidGuid(args.options.userId)) {
        return `${args.options.userId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently in preview
    and is subject to change once the API reached general availability.

    The ${chalk.grey(`appId`)} has to be the id of the app instance installed for the user.
    Do not use the ID from the manifest of the zip app package or the id
    from the Microsoft Teams App Catalog.

  Examples:

    Uninstall an app for the specified user
      ${this.name} --appId YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY= --userId 2609af39-7775-4f94-a3dc-0dd67657e900
`);
  }
}

module.exports = new TeamsUserAppRemoveCommand();