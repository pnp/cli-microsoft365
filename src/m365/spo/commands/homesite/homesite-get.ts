import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { Logger } from '../../../../cli/Logger.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';

const options = globalOptionsZod
  .extend({
    url: zod.alias('u', z.string()
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

class SpoHomeSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_GET;
  }

  public get description(): string {
    return 'Gets information about a home site';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.verbose);

      if (this.verbose) {
        await logger.log(`Retrieving home sites...`);
      }
      const homeSites = await odata.getAllItems<{ Url: string; }>(`${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`);

      const homeSite = homeSites.find(hs => urlUtil.removeTrailingSlashes(hs.Url).toLowerCase() === urlUtil.removeTrailingSlashes(args.options.url).toLowerCase());

      if (homeSite === undefined) {
        throw `Home site with URL '${args.options.url}' not found.`;
      }

      await logger.log(homeSite);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHomeSiteGetCommand();