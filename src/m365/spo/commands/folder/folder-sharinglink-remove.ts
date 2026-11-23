import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../../../../request.js';
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
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_REMOVE;
  }

  public get description(): string {
    return 'Removes a sharing link from a folder';
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
    const removeSharingLink = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing sharing link of folder ${args.options.folderId || args.options.folderUrl} with id ${args.options.id}...`);
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
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeSharingLink();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove sharing link ${args.options.id} of folder ${args.options.folderUrl || args.options.folderId}?` });

      if (result) {
        await removeSharingLink();
      }
    }
  }
}

export default new SpoFolderSharingLinkRemoveCommand();
