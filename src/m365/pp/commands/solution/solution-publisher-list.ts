import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { Publisher } from './Solution.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  includeMicrosoftPublishers?: boolean;
  withMicrosoftPublishers?: boolean;
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
        withMicrosoftPublishers: typeof args.options.withMicrosoftPublishers !== 'undefined',
        asAdmin: !!args.options.asAdmin
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--includeMicrosoftPublishers'
      },
      {
        option: '--withMicrosoftPublishers'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.includeMicrosoftPublishers) {
      await this.warn(logger, `Parameter 'includeMicrosoftPublishers' is deprecated. Please use 'withMicrosoftPublishers' instead`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of publishers...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const shouldIncludeMicrosoftPublishers: boolean | undefined = args.options.withMicrosoftPublishers || args.options.includeMicrosoftPublishers;
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.0/publishers?$select=publisherid,uniquename,friendlyname,versionnumber,isreadonly,description,customizationprefix,customizationoptionvalueprefix${!shouldIncludeMicrosoftPublishers ? `&$filter=publisherid ne 'd21aab70-79e7-11dd-8874-00188b01e34f'` : ''}&api-version=9.1`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: Publisher[] }>(requestOptions);
      await logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpSolutionPublisherListCommand();