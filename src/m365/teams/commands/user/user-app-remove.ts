import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandValidate } from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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
            cmd.log(chalk.green('DONE'));
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
      if (!Utils.isValidGuid(args.options.userId)) {
        return `${args.options.userId} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new TeamsUserAppRemoveCommand();