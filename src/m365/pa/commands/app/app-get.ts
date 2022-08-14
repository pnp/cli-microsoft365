import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import PowerAppsCommand from '../../../base/PowerAppsCommand';
import commands from '../../commands';
import * as paAppListCommand from '../app/app-list';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  displayName?: string;
}

class PaAppGetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power App';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName', 'description', 'appVersion', 'owner'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        displayName: typeof args.options.displayName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '-d, --displayName [displayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {    
        if (args.options.name && !validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['name', 'displayName']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.name) {
      const requestOptions: any = {
        url: `${this.resource}/providers/Microsoft.PowerApps/apps/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
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
}

module.exports = new PaAppGetCommand();