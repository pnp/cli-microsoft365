import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { pa } from '../../../../utils/pa.js';
import { validation } from '../../../../utils/validation.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

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
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/providers/Microsoft.PowerApps/apps/${formatting.encodeQueryParameter(args.options.name)}?api-version=2016-11-01`,
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

        const app = await pa.getAppByDisplayName(args.options.displayName!);

        if (!!app) {
          logger.log(this.setProperties(app));
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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