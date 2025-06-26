import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { zod } from '../../../../utils/zod.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { formatting } from '../../../../utils/formatting.js';
import { Page } from './Page.js';

const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string())
      .refine(url => validation.isValidSharePointUrl(url) === true, url => ({
        message: `'${url}' is not a valid SharePoint URL.`
      })),
    pageName: zod.alias('n', z.string()),
    id: zod.alias('i', z.string())
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })),
    draft: z.boolean().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoPageControlRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_CONTROL_REMOVE;
  }

  public get description(): string {
    return 'Removes a control from a modern page';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.force) {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to delete control '${args.options.id}' on page '${args.options.pageName}'?` });

      if (!result) {
        return;
      }
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Getting page properties for page '${args.options.pageName}'...`);
      }

      const pageName = args.options.pageName.toLowerCase().endsWith('.aspx') ? args.options.pageName : `${args.options.pageName}.aspx`;
      const serverRelativePageUrl = urlUtil.getServerRelativePath(args.options.webUrl, `SitePages/${urlUtil.removeLeadingSlashes(pageName)}`);

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(serverRelativePageUrl)}')?$select=UniqueId,SiteId`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const pageProperties = await request.get<{ UniqueId: string; SiteId: string }>(requestOptions);

      if (this.verbose) {
        await logger.logToStderr(`Removing control '${args.options.id}' on page '${args.options.pageName}'...`);
      }

      const sharePointRootUrl = new URL(args.options.webUrl).origin;
      requestOptions = {
        url: `${sharePointRootUrl}/_api/v2.1/sites/${pageProperties.SiteId}/pages/${pageProperties.UniqueId}/oneDrive.page/webParts/${args.options.id}`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);

      if (!args.options.draft) {
        if (this.verbose) {
          await logger.logToStderr(`Republishing page...`);
        }

        await Page.publishPage(args.options.webUrl, pageName);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoPageControlRemoveCommand();