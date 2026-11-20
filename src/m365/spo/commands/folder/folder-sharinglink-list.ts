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
  }).optional().alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkListCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_LIST;
  }

  public get description(): string {
    return 'Lists all the sharing links of a specific folder';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'scope', 'roles', 'link'];
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
      await logger.logToStderr(`Retrieving sharing links for folder ${args.options.folderId || args.options.folderUrl}...`);
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

      // remove grantedToIdentities from the sharing link object
      const filteredSharingLinks = sharingLinks.map(link => {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const { grantedToIdentities, ...filteredLink } = link;
        return filteredLink;
      });

      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        await logger.log(filteredSharingLinks);
      }
      else {
        //converted to text friendly output
        await logger.log(filteredSharingLinks.map(i => {
          return {
            id: i.id,
            roles: i.roles.join(','),
            link: i.link.webUrl,
            scope: i.link.scope
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkListCommand();
