import { cli, CommandOutput } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';
import paAppListCommand from '../app/app-list.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  displayName?: string;
  environmentName?: string;
  asAdmin?: boolean;
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
        displayName: typeof args.options.displayName !== 'undefined',
        asAdmin: !!args.options.asAdmin,
        environmentName: typeof args.options.environmentName !== 'undefined'
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
      },
      {
        option: '-e, --environmentName [environmentName]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.name && !validation.isValidGuid(args.options.name)) {
          return `${args.options.name} is not a valid GUID`;
        }

        if (args.options.asAdmin && !args.options.environmentName) {
          return 'When specifying the asAdmin option, the environment option is required as well.';
        }

        if (args.options.environmentName && !args.options.asAdmin) {
          return 'When specifying the environment option, the asAdmin option is required as well.';
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
        let endpoint = `${this.resource}/providers/Microsoft.PowerApps`;
        if (args.options.asAdmin) {
          endpoint += `/scopes/admin/environments/${formatting.encodeQueryParameter(args.options.environmentName!)}`;
        }
        endpoint += `/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`;

        const requestOptions: CliRequestOptions = {
          url: endpoint,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        if (this.verbose) {
          await logger.logToStderr(`Retrieving information about Microsoft Power App with name '${args.options.name}'...`);
        }

        const res = await request.get<any>(requestOptions);
        await logger.log(this.setProperties(res));
      }
      else {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving information about Microsoft Power App with displayName '${args.options.displayName}'...`);
        }

        const getAppsOutput = await this.getApps(args, logger);
        if (getAppsOutput.stdout && JSON.parse(getAppsOutput.stdout).length > 0) {
          const allApps: any[] = JSON.parse(getAppsOutput.stdout);
          const app = allApps.find((a: any) => {
            return a.properties.displayName.toLowerCase() === `${args.options.displayName}`.toLowerCase();
          });

          if (!!app) {
            await logger.log(this.setProperties(app));
          }
          else {
            throw `No app found with displayName '${args.options.displayName}'.`;
          }
        }
        else {
          throw 'No apps found.';
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getApps(args: CommandArgs, logger: Logger): Promise<CommandOutput> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving all apps...`);
    }

    const options: GlobalOptions = {
      output: 'json',
      debug: this.debug,
      verbose: this.verbose,
      environmentName: args.options.environmentName,
      asAdmin: args.options.asAdmin
    };

    return await cli.executeCommandWithOutput(paAppListCommand as Command, { options: { ...options, _: [] } });
  }

  private setProperties(app: any): any {
    app.displayName = app.properties.displayName;
    app.description = app.properties.description || '';
    app.appVersion = app.properties.appVersion;
    app.owner = app.properties.owner.email || '';
    return app;
  }
}

export default new PaAppGetCommand();