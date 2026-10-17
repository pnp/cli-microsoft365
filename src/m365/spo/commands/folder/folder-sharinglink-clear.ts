import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../../../../request.js';
import commands from '../../commands.js';

const allowedScopes = ['anonymous', 'users', 'organization'] as const;

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
  scope: z.enum(allowedScopes, {
    error: e => `'${e.input}' is not a valid scope. Allowed values are: ${allowedScopes.join(', ')}`
  }).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkClearCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_CLEAR;
  }

  public get description(): string {
    return 'Removes all sharing links of a folder';
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
    const clearSharingLinks = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Clearing sharing links from folder ${args.options.folderId || args.options.folderUrl} for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}`);
      }

      try {
        const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
        const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
        const folderUrl: URL = new URL(absoluteFolderUrl);

        const siteId: string = await spo.getSiteIdByMSGraph(args.options.webUrl);
        const driveDetails: Drive = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
        const itemId: string = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);

        let requestUrl = `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions?$filter=Link ne null`;
        if (args.options.scope) {
          requestUrl += ` and Link/Scope eq '${args.options.scope}'`;
        }

        const sharingLinks = await odata.getAllItems<any>(requestUrl);

        for (const sharingLink of sharingLinks) {
          const requestOptions: CliRequestOptions = {
            url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/permissions/${sharingLink.id}`,
            headers: {
              accept: 'application/json;odata.metadata=none'
            }
          };
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await clearSharingLinks();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to clear sharing links from folder ${args.options.folderUrl || args.options.folderId}? for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}` });

      if (result) {
        await clearSharingLinks();
      }
    }
  }

}

export default new SpoFolderSharingLinkClearCommand();