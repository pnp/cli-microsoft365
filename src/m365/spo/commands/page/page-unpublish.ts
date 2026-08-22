import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string().alias('n')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageUnpublishCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_UNPUBLISH;
  }

  public get description(): string {
    return 'Unpublishes a modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let pageName: string = urlUtil.removeLeadingSlashes(args.options.name);
      if (!pageName.toLowerCase().endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        await logger.logToStderr(`Unpublishing page ${pageName}...`);
      }

      const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')/UnPublish`,
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

export default new SpoPageUnpublishCommand();
