import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';
import * as paAppListCommand from '../app/app-list';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  displayName?: string;
}

class PaAppGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power App';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName', 'description', 'appVersion', 'owner'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.name) {
      let requestUrl: string = '';
      requestUrl = `${this.resource}providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`;

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      if (this.verbose) {
        logger.logToStderr(`Retrieving information about Microsoft Power App with name '${args.options.name}'...`);
      }

      request
        .get(requestOptions)
        .then((res: any): void => {
          logger.log(this.setProperties(res));
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
    else {
      if (this.verbose) {
        logger.logToStderr(`Retrieving information about Microsoft Power App with displayName '${args.options.displayName}'...`);
      }

      this
        .getApps(args, logger)
        .then((getAppsOutput: CommandOutput): void => {
          const allApps: any = JSON.parse(getAppsOutput.stdout);
          if (allApps.length > 0) {
            const app = allApps.find((a: any) => {
              return a.properties.displayName.toLowerCase() === `${args.options.displayName}`.toLowerCase();
            });
            if (!!app) {
              logger.log(this.setProperties(app));
            }
            else {
              if (this.verbose) {
                logger.logToStderr(`No app found with displayName '${args.options.displayName}'`);
              }
            }
          }
          else {
            if (this.verbose) {
              logger.logToStderr('No apps found');
            }
          }
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    }
  }

  private getApps(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all apps...`);
    }

    const options: GlobalOptions = {
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(paAppListCommand as Command, { options: { ...options, _: [] } });
  }

  private setProperties(app: any) {
    app.displayName = app.properties.displayName;
    app.description = app.properties.description || '';
    app.appVersion = app.properties.appVersion;
    app.owner = app.properties.owner.email || '';
    return app;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name [name]'
      },
      {
        option: '-d, --displayName [displayName]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.name && !args.options.displayName) {
      return 'Specify either name or displayName';
    }

    if (args.options.name && args.options.displayName) {
      return 'Specify either name or displayName but not both';
    }

    if (args.options.name && !validation.isValidGuid(args.options.name)) {
      return `${args.options.name} is not a valid GUID`;
    }

    return true;
  }

}

module.exports = new PaAppGetCommand();
