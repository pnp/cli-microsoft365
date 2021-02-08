import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeTheme = (): void => {
      this
        .getSpoAdminUrl(logger, this.debug)
        .then((spoAdminUrl: string): Promise<void> => {
          if (this.verbose) {
            logger.logToStderr(`Removing theme from tenant...`);
          }

          const requestOptions: any = {
            url: `${spoAdminUrl}/_api/thememanager/DeleteTenantTheme`,
            headers: {
              'accept': 'application/json;odata=nometadata'
            },
            data: {
              name: args.options.name,
            },
            responseType: 'json'
          };

          return request.post(requestOptions);
        })
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    }

    if (args.options.confirm) {
      removeTheme();
    }
    else {
      Cli.prompt({
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
        option: '-n, --name <name>'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new SpoThemeRemoveCommand();