import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { powerPlatform } from '../../../../utils/powerPlatform.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Webrole } from './Webrole.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  websiteId: z.uuid().optional(),
  websiteName: z.string().optional(),
  environmentName: z.string().alias('e'),
  asAdmin: z.boolean().optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpWebSiteWebRoleListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.WEBSITE_WEBROLE_LIST;
  }

  public get description(): string {
    return 'Lists all webroles for the specified Power Pages website.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.websiteId, options.websiteName].filter(x => x !== undefined).length === 1, {
        error: `Specify either websiteId or websiteName, but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the webroles for '${args.options.websiteId || args.options.websiteName}'...`);
    }

    try {
      const dynamicsApiUrl = await powerPlatform.getDynamicsInstanceApiUrl(args.options.environmentName, args.options.asAdmin);
      const websiteRecordId = await this.getWebsiteRecordId(args, dynamicsApiUrl);
      const roles = await this.getWebsiteRoles(dynamicsApiUrl, websiteRecordId);

      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        await logger.log(roles);
      }
      else {
        // converted to text friendly output
        await logger.log(roles.map(i => {
          return {
            webroleid: i.mspp_webroleid,
            name: i.mspp_name,
            statuscode: i.statuscode
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getWebsiteRecordId(args: CommandArgs, dynamicsApiUrl: string): Promise<string> {
    if (args.options.websiteId) {
      const website = await powerPlatform.getWebsiteById(args.options.environmentName, args.options.websiteId);
      return website.websiteRecordId;
    }
    return powerPlatform.getWebsiteIdByUniqueName(dynamicsApiUrl, args.options.websiteName!);
  }

  private async getWebsiteRoles(dynamicsApiUrl: string, websiteId: string): Promise<Webrole[]> {
    const requestUrl = `${dynamicsApiUrl}/api/data/v9.2/mspp_webroles?$filter=_mspp_websiteid_value eq ${websiteId}`;
    const result = await odata.getAllItems<Webrole>(requestUrl);
    return result;
  }
}

export default new PpWebSiteWebRoleListCommand();