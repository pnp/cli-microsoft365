import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
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
  name: z.string().alias('n'),
  permanent: z.boolean().optional(),
  bypassSharedLock: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a modern page';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removePage(logger, args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove page '${args.options.name}'?` });

      if (result) {
        await this.removePage(logger, args);
      }
    }
  }

  private async removePage(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      // Remove leading slashes from the page name (page can be nested in folders)
      let pageName: string = urlUtil.removeLeadingSlashes(args.options.name);
      if (!pageName.toLowerCase().endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        await logger.logToStderr(`Removing page ${pageName}...`);
      }

      const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.bypassSharedLock) {
        requestOptions.headers!.Prefer = 'bypass-shared-lock';
      }
      if (args.options.permanent) {
        await request.delete(requestOptions);
      }
      else {
        requestOptions.url += '/Recycle';
        await request.post(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageRemoveCommand();
