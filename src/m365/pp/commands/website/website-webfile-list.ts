import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { PpWebSiteOptions } from '../Website.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  websiteId?: string;
  websiteName?: string;
  asAdmin?: boolean;
}

class PpWebSiteWebFileListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.WEBSITE_WEBFILE_LIST;
  }

  public get description(): string {
    return 'List all webfiles for the specified Power Pages website';
  }

  public defaultProperties(): string[] | undefined {
    return ['mspp_name', 'mspp_webfileid', 'mspp_summary', '_mspp_publishingstateid_value@OData.Community.Display.V1.FormattedValue'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionsSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        websiteId: typeof args.options.websiteId !== 'undefined',
        websiteName: typeof args.options.name !== 'undefined',
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
        option: '--websiteId [websiteId]'
      },
      {
        option: '--websiteName [websiteName]'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initOptionsSets(): void {
    this.optionSets.push(
      { options: ['websiteId', 'websiteName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.websiteId && !validation.isValidGuid(args.options.websiteId)) {
          return `${args.options.websiteId} is not a valid GUID`;
        }
        return true;
      }
    );
  }

  private async getWebSiteId(dynamicsApiUrl: string, args: CommandArgs): Promise<any> {
    if (args.options.websiteId) {
      return args.options.websiteId;
    }
    const options: PpWebSiteOptions = {
      environmentName: args.options.environmentName,
      id: args.options.websiteId,
      name: args.options.websiteName,
      output: 'json'
    };

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'applciation/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (options.name) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/powerpagesites?$filter=name eq '${options.name}'&$select=powerpagesiteid`;
      const result = await request.get<{ value: any[] }>(requestOptions);

      if (result.value.length === 0) {
        throw `The specified website '${args.options.websiteName}' does not exist.`;
      }
      return result.value[0].powerpagesiteid;
    }
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of webfiles`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const websiteId = await this.getWebSiteId(dynamicsApiUrl, args);

      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.2/mspp_webfiles?$filter=_mspp_websiteid_value eq '${websiteId}'`,
        headers: {
          accept: `application/json;`,
          'odata-version': '4.0',
          prefer: `odata.include-annotations="*"`
        },
        responseType: 'json'
      };

      const items = await odata.getAllItems<any>(requestOptions);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpWebSiteWebFileListCommand();