import { z } from 'zod';
import { Permission } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import commands from '../../commands.js';
import { drive as driveItem } from '../../../../utils/drive.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  folderUrl: z.string().optional(),
  folderId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option 'folderId'.`
    })
    .optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option 'id'.`
    })
    .alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_GET;
  }

  public get description(): string {
    return 'Gets details about a specific sharing link on a folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.folderUrl, opts.folderId].filter(o => o !== undefined).length === 1, {
        error: `Specify one of the following options: folderUrl, folderId.`,
        params: {
          customCode: 'optionSet',
          options: ['folderUrl', 'folderId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving sharing link on folder ${args.options.folderId || args.options.folderUrl}...`);
    }

    try {
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteIdByMSGraph(args.options.webUrl);
      const drive: Drive = await driveItem.getDriveByUrl(siteId, folderUrl);
      const itemId: string = await driveItem.getDriveItemId(drive, folderUrl);

      const requestUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${itemId}/permissions/${args.options.id}`;

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const permission = await request.get<Permission>(requestOptions);
      await logger.log(permission);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkGetCommand();