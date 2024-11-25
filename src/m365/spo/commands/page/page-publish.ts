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
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint Online site URL.`
      }))
    ),
    name: zod.alias('n', z.string())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPagePublishCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_PUBLISH;
  }

  public get description(): string {
    return 'Publishes a modern page';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      // Remove leading slashes from the page name (page can be nested in folders)
      let pageName: string = urlUtil.removeLeadingSlashes(args.options.name);
      if (!pageName.toLowerCase().endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        await logger.logToStderr(`Publishing page ${pageName}...`);
      }

      const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')/Publish()`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPagePublishCommand();
