import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  confirm?: boolean;
}

class PaAppRemoveCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified Microsoft Power App';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Removing Microsoft Power App ${args.options.name}...`);
    }

    const removePaApp: () => void = (): void => {
      const requestOptions: any = {
        url: `${this.resource}/providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2017-08-01`,
        resolveWithFullResponse: true,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      request
        .delete(requestOptions)
        .then((): void => cb(), (rawRes: any): void => {
          if (rawRes.response && rawRes.response.status === 403) {
            cb(`App '${args.options.name}' does not exist`);
          }
          else {
            this.handleRejectedODataJsonPromise(rawRes, logger, cb);
          }
        });
    };

    if (args.options.confirm) {
      removePaApp();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the Microsoft Power App ${args.options.name}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removePaApp();
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

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.name)) {
      return `${args.options.name} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new PaAppRemoveCommand();