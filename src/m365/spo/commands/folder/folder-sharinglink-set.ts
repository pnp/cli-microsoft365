import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import commands from '../../commands.js';

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
    .alias('i'),
  expirationDateTime: z.string()
    .refine(dateTime => validation.isValidISODateTime(dateTime), {
      error: e => `'${e.input}' is not a valid ISO date string.`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_SET;
  }

  public get description(): string {
    return 'Updates a specific sharing link to a folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.folderId, opts.folderUrl].filter(o => o !== undefined).length === 1, {
        error: `Specify one of the following options: folderId, folderUrl.`,
        params: {
          customCode: 'optionSet',
          options: ['folderId', 'folderUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
    }

    try {
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteIdByMSGraph(args.options.webUrl);
      const driveDetails: Drive = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
      const itemId: string = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          expirationDateTime: args.options.expirationDateTime
        }
      };

      const sharingLink = await request.patch<any>(requestOptions);

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkSetCommand();