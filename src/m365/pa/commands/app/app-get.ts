import { Cli } from '../../../../cli/Cli';
import { CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
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
    this.optionSets.push({ options: ['name', 'displayName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.name) {
        const requestOptions: any = {
          url: `${this.resource}/providers/Microsoft.PowerApps/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        if (this.verbose) {
          logger.logToStderr(`Retrieving information about Microsoft Power App with name '${args.options.name}'...`);
        }

        const res = await request.get<any>(requestOptions);
        logger.log(this.setProperties(res));
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Retrieving information about Microsoft Power App with displayName '${args.options.displayName}'...`);
        }

        const getAppsOutput = await this.getApps(args, logger);

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
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
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

  private setProperties(app: any): any {
    app.displayName = app.properties.displayName;
    app.description = app.properties.description || '';
    app.appVersion = app.properties.appVersion;
    app.owner = app.properties.owner.email || '';
    return app;
  }
}

module.exports = new PaAppGetCommand();