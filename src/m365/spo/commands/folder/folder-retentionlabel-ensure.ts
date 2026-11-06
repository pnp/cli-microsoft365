import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FolderProperties } from './FolderProperties.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  name: z.string(),
  folderUrl: z.string().optional(),
  folderId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID.`
    })
    .optional()
    .alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Applies a retention label to a folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.folderUrl, opts.folderId].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'folderUrl' or 'folderId', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['folderUrl', 'folderId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const folderProperties = await this.getFolderProperties(logger, args);

      if (folderProperties.ListItemAllFields) {
        const parsedUrl = new URL(args.options.webUrl);
        const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
        const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, folderProperties.ListItemAllFields.ParentList.RootFolder.ServerRelativeUrl);

        await spo.applyRetentionLabelToListItems(args.options.webUrl, args.options.name, listAbsoluteUrl, [parseInt(folderProperties.ListItemAllFields.Id)], logger, args.options.verbose);
      }
      else {
        const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, folderProperties.ServerRelativeUrl);

        await spo.applyDefaultRetentionLabelToList(args.options.webUrl, args.options.name, listAbsoluteUrl, false, logger, args.options.verbose);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderProperties(logger: Logger, args: CommandArgs): Promise<FolderProperties> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list and item information for folder '${args.options.folderId || args.options.folderUrl}' in site at ${args.options.webUrl}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.folderId) {
      requestUrl += `GetFolderById('${args.options.folderId}')`;
    }
    else {
      const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl!);
      requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=ListItemAllFields,ListItemAllFields/ParentList/RootFolder&$select=ServerRelativeUrl,ListItemAllFields/ParentList/RootFolder/ServerRelativeUrl,ListItemAllFields/Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return await request.get<FolderProperties>(requestOptions);
  }
}

export default new SpoFolderRetentionLabelEnsureCommand();