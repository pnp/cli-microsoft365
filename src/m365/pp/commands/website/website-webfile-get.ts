import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { PpWebSiteOptions } from '../Website.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  environmentName: string;
  id?: string;
  websiteId?: string;
  websiteName?: string;
  asAdmin?: boolean;
}

class PpWebSiteWebFileGetCommand extends PowerPlatformCommand {

  public get name(): string {
    return commands.WEBSITE_WEBFILE_GET;
  }

  public get description(): string {
    return 'Gets information about the specified web file';
  }

  public defaultProperties(): string[] | undefined {
    return ['mspp_name', 'mspp_webfileid', 'mspp_summary', '_mspp_publishingstateid_value@OData.Community.Display.V1.FormattedValue'];
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
        id: typeof args.options.id !== 'undefined',
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
        option: '-i, --id [id]'
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

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['websiteId', 'websiteName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }
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
      await logger.logToStderr(`Retrieving a website webfile '${args.options.websiteId || args.options.websiteName}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);

      const websiteId = await this.getWebSiteId(dynamicsApiUrl, args);

      const res = await this.getWebSiteWebFile(dynamicsApiUrl, websiteId, args.options);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getWebSiteWebFile(dynamicsApiUrl: string, websiteId: string, options: Options): Promise<any> {

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none',
        'odata-version': '4.0',
        prefer: 'odata.include-annotations="*"'
      },
      responseType: 'json'
    };

    // let webfileitem: any | null = null;
    // requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/mspp_webfiles(${options.id})`;
    // const result = await request.get<any>(requestOptions);
    // webfileitem = result;

    // if (!webfileitem || webfileitem["_mspp_websiteid_value"] !== websiteId) {
    //   throw `The specified webfile '${options.id}' does not exist on website.`;
    // }

    // return webfileitem;


    requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/mspp_webfiles?$filter=mspp_webfileid eq '${options.id}' and _mspp_websiteid_value eq '${websiteId}'`;
    const result = await request.get<{ value: any[] }>(requestOptions);

    if (result.value.length === 0) {
      throw `The specified webfile '${options.id}' does not exist.`;
    }

    return result.value[0];




    // requestOptions.url = `${dynamicsApiUrl}/api/data/v9.1/websitewebfiles?$filter=name eq '${options.name}'`;
    // const result = await request.get<{ value: any[] }>(requestOptions);

    // if (result.value.length === 0) {
    //   throw `The specified websitewebfile '${options.name}' does not exist.`;
    // }

    // if (result.value.length > 1) {
    //   const resultAsKeyValuePair = formatting.convertArrayToHashTable('websitewebfileid', result.value);
    //   return cli.handleMultipleResultsFound(`Multiple websitewebfiles with name '${options.name}' found`, resultAsKeyValuePair);
    // }

    // return result.value[0];

  }
}

export default new PpWebSiteWebFileGetCommand();