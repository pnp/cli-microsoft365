import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { odata } from '../../../../utils/odata.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import { validation } from '../../../../utils/validation.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { PpWebSiteOptions } from '../Website.js';



export const options = globalOptionsZod
  .extend({
    environmentName: zod.alias('e', z.string()),
    websiteId: z.string().refine(id => validation.isValidGuid(id) === true, id => ({ message: `${id} is not a valid GUID` })).optional(),
    websiteName: z.string().optional(),
    asAdmin: z.boolean().optional()
  })
  .refine(options => !(options.websiteId !== undefined && options.websiteName !== undefined), {
    message: `Either websiteId or websiteName is required, but not both.`
  })
  .refine(options => !(options.websiteId === undefined && options.websiteName === undefined), {
    message: `Either websiteId or websiteName is required.`
  });

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpWebSiteWebLinkListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.WEBSITE_WEBLINK_LIST;
  }

  public get description(): string {
    return 'List all weblinks for the specified Power Pages website';
  }

  public defaultProperties(): string[] | undefined {
    return ['mspp_name', 'mspp_weblinkid', 'mspp_description', 'statecode'];
  }

  public get schema(): z.ZodTypeAny {
    return options;
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
        accept: 'application/json;odata.metadata=none'
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

  private async getwebsitelinksets(dynamicsApiUrl: string, websiteId: string): Promise<any> {

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (websiteId) {
      requestOptions.url = `${dynamicsApiUrl}/api/data/v9.2/mspp_weblinksets?$filter=_mspp_websiteid_value eq '${websiteId}'`;
      const result = await request.get<{ value: any[] }>(requestOptions);
      if (result.value.length === 0) {
        throw `The specified website '${websiteId}' does not have links.`;
      }
      const weblinksets = result.value.map(linkset => "'" + linkset.mspp_weblinksetid.toString() + "'");
      return weblinksets.join(',');
    }
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of weblinks from website ${args.options.websiteId || args.options.websiteName} in environment ${args.options.environmentName}...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const websiteId = await this.getWebSiteId(dynamicsApiUrl, args);
      const weblinksets = await this.getwebsitelinksets(dynamicsApiUrl, websiteId);
      const requestOptions: CliRequestOptions = {
        url: `${dynamicsApiUrl}/api/data/v9.2/mspp_weblinks?$filter=Microsoft.Dynamics.CRM.ContainValues(PropertyName=@p1,PropertyValues=@p2)&@p1='mspp_weblinksetid'&@p2=[${weblinksets}]`,
        headers: {
          accept: `application/json;`
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

export default new PpWebSiteWebLinkListCommand();