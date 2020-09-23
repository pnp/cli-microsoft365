import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId: string;
  teamId: string;
  confirm?: boolean;
}

class TeamsAppUninstallCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_UNINSTALL}`;
  }

  public get description(): string {
    return 'Uninstalls an app from a Microsoft Team team';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const uninstallApp: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/teams/${args.options.teamId}/installedApps/${args.options.appId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }

          cb();
        }, (res: Error): void => this.handleRejectedODataJsonPromise(res, logger, cb));
    };

    if (args.options.confirm) {
      uninstallApp();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to uninstall the app with id ${args.options.appId} from the Microsoft Teams team ${args.options.teamId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          uninstallApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--appId <appId>',
        description: 'The unique id of the app instance installed in the Microsoft Teams team'
      },
      {
        option: '--teamId <teamId>',
        description: 'The ID of the Microsoft Teams team from which to uninstall the app'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation when uninstalling the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId)) {
      return `${args.options.teamId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new TeamsAppUninstallCommand();