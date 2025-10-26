import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = globalOptionsZod
  .extend({
    siteUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    )
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteVersionpolicyGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_VERSIONPOLICY_GET;
  }

  public get description(): string {
    return 'Retrieves the version policy settings of a specific site';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving version policy settings for site '${args.options.siteUrl}'...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/site/VersionPolicyForNewLibrariesTemplate?$expand=VersionPolicies`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.get<any>(requestOptions);

      let defaultTrimMode: string = 'number';
      if (response.MajorVersionLimit === -1) {
        defaultTrimMode = 'inheritTenant';
      }
      else if (response.VersionPolicies) {
        switch (response.VersionPolicies.DefaultTrimMode) {
          case 1:
            defaultTrimMode = 'age';
            break;
          case 2:
            defaultTrimMode = 'automatic';
            break;
          case 0:
          default:
            defaultTrimMode = 'number';
        }
      }

      const output = {
        defaultTrimMode: defaultTrimMode,
        defaultExpireAfterDays: response.VersionPolicies?.DefaultExpireAfterDays ?? null,
        majorVersionLimit: response.MajorVersionLimit
      };

      await logger.log(output);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteVersionpolicyGetCommand();