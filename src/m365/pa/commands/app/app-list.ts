import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment?: string;
  asAdmin: boolean;
}

class PaAppListCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists all Power Apps apps';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName'];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.asAdmin = args.options.asAdmin === true;
    telemetryProps.environment = typeof args.options.environment !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const url = `${this.resource}/providers/Microsoft.PowerApps${args.options.asAdmin ? '/scopes/admin' : ''}${args.options.environment ? '/environments/' + encodeURIComponent(args.options.environment) : ''}/apps?api-version=2017-08-01`;

    odata
      .getAllItems<{ name: string; displayName: string; properties: { displayName: string } }>(url)
      .then((apps: { name: string; displayName: string; properties: { displayName: string } }[]): void => {
        if (apps.length > 0) {
          apps.forEach(a => {
            a.displayName = a.properties.displayName;
          });

          logger.log(apps);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No apps found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --environment [environment]'
      },
      {
        option: '--asAdmin'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.asAdmin && !args.options.environment) {
      return 'When specifying the asAdmin option the environment option is required as well';
    }

    if (args.options.environment && !args.options.asAdmin) {
      return 'When specifying the environment option the asAdmin option is required as well';
    }

    return true;
  }

}

module.exports = new PaAppListCommand();