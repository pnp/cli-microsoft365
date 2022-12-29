import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { powerPlatform } from '../../../../utils/powerPlatform';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';
import { Publisher } from './Solution';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  includeMicrosoftPublishers: boolean;
  asAdmin: boolean;
}

class PpSolutionPublisherListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.SOLUTION_PUBLISHER_LIST;
  }

  public get description(): string {
    return 'Lists publishers in a given environment.';
  }

  public defaultProperties(): string[] | undefined {
    return ['publisherid', 'uniquename', 'friendlyname'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        includeMicrosoftPublishers: typeof args.options.includeMicrosoftPublishers !== 'undefined',
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--includeMicrosoftPublishers'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of publishers...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environment, args.options.asAdmin);

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix${!args.options.includeMicrosoftPublishers ? `&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'` : ''}&api-version=9.1`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Publisher[] }>(requestOptions);
      logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpSolutionPublisherListCommand();