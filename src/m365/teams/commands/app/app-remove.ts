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
  confirm?: boolean;
  id: string;
}

class TeamsAppRemoveCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_APP_REMOVE}`;
  }

  public get description(): string {
    return 'Removes a Teams app from the organization\'s app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const { id: appId } = args.options;

    const removeApp: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/v1.0/appCatalogs/teamsApps/${appId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      };

      request
        .delete(requestOptions)
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (res: any): void => this.handleRejectedODataJsonPromise(res, cmd, cb));
    };

    if (args.options.confirm) {
      removeApp();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Teams app ${appId} from the app catalog?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeApp();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'ID of the Teams app to remove. Needs to be available in your organization\'s app catalog.'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the app'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }
}

module.exports = new TeamsAppRemoveCommand();