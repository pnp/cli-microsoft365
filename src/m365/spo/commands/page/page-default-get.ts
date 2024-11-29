import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
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

class SpoPageDefaultGetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_DEFAULT_GET;
  }

  public get description(): string {
    return 'Gets the home page for a specific site';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving home page information for site: '${args.options.siteUrl}'...`);
      }

      let requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/Web/RootFolder?$select=WelcomePage`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const { WelcomePage } = await request.get<{ WelcomePage: string }>(requestOptions);

      if (this.verbose) {
        await logger.logToStderr(`Home page URL retrieved: '${WelcomePage}'. Fetching the details...`);
      }

      requestOptions = {
        url: `${args.options.siteUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${urlUtil.getServerRelativeSiteUrl(args.options.siteUrl)}/${formatting.encodeQueryParameter(WelcomePage)}')?$expand=ListItemAllFields/ClientSideApplicationId,ListItemAllFields/PageLayoutType,ListItemAllFields/CommentsDisabled`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const page = await request.get<any>(requestOptions);

      await logger.log(page);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageDefaultGetCommand();