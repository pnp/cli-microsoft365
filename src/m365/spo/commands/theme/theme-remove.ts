import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  confirm?: boolean;
}

class SpoThemeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.THEME_REMOVE;
  }

  public get description(): string {
    return 'Removes existing theme';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeTheme = (): void => {
      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((spoAdminUrl: string): Promise<void> => {
          if (this.verbose) {
            cmd.log(`Removing theme from tenant...`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_api/thememanager/DeleteTenantTheme`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            body: {
              name: args.options.name,
            },
            json: true
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
    }

    if (args.options.confirm) {
      removeTheme();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the theme`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeTheme();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Name of the theme to remove'
      },
      {
        option: '--confirm',
        description: 'Do not prompt for confirmation before removing theme'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoThemeRemoveCommand();