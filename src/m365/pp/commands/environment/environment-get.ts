import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name?: string;
  asAdmin?: boolean;
}

class PpEnvironmentGetCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Power Platform environment';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name [name]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving environment: ${args.options.name || 'default'}`);
    }

    let url: string = `${this.resource}/providers/Microsoft.BusinessAppPlatform`;
    if (args.options.asAdmin) {
      url += '/scopes/admin';
    }

    const envName = args.options.name ? formatting.encodeQueryParameter(args.options.name) : '~Default';
    url += `/environments/${envName}?api-version=2020-10-01`;

    const requestOptions: CliRequestOptions = {
      url: url,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<any>(requestOptions);
    await logger.log(response);
  }
}

export default new PpEnvironmentGetCommand();