import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerAppsCommand from '../../../base/PowerAppsCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
}

class PaEnvironmentGetCommand extends PowerAppsCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Power Apps environment';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Power Apps environment ${args.options.name || 'default'}...`);
    }

    const environmentName = args.options.name ? formatting.encodeQueryParameter(args.options.name) : '~default';
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.PowerApps/environments/${environmentName}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);

      res.displayName = res.properties.displayName;
      res.provisioningState = res.properties.provisioningState;
      res.environmentSku = res.properties.environmentSku;
      res.azureRegionHint = res.properties.azureRegionHint;
      res.isDefault = res.properties.isDefault;

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PaEnvironmentGetCommand();