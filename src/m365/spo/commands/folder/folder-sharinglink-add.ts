import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import commands from '../../commands.js';

const allowedTypes = ['view', 'edit'] as const;
const allowedScopes = ['anonymous', 'organization', 'users'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  folderId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option 'folderId'.`
    })
    .optional(),
  folderUrl: z.string().optional(),
  type: z.string().pipe(z.enum(allowedTypes, {
    error: e => `'${e.input}' is not a valid type. Allowed values are: ${allowedTypes.join(',')}`
  })),
  expirationDateTime: z.string()
    .refine(dateTime => validation.isValidISODateTime(dateTime), {
      error: e => `'${e.input}' is not a valid ISO date string.`
    })
    .optional(),
  scope: z.enum(allowedScopes, {
    error: e => `'${e.input}' is not a valid scope. Allowed values are: ${allowedScopes.join(', ')}.`
  }).optional(),
  retainInheritedPermissions: z.boolean().optional(),
  recipients: z.string()
    .refine(recipients => validation.isValidUserPrincipalNameArray(recipients) === true, {
      error: e => {
        const invalidUpns = String(e.input).split(',').map(upn => upn.trim()).filter(upn => !validation.isValidUserPrincipalName(upn));
        return `The following user principal names are invalid for the option 'recipients': ${invalidUpns.join(', ')}.`;
      }
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderSharingLinkAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_ADD;
  }

  public get description(): string {
    return 'Creates a new sharing link to a folder';
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
      })
      .refine(opts => opts.scope !== 'users' || opts.recipients !== undefined, {
        error: `Specify recipients when the scope is set to 'users'.`,
        path: ['recipients'],
        params: {
          customCode: 'required'
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating a sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
    }

    try {
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteIdByMSGraph(args.options.webUrl);
      const driveDetails: Drive = await drive.getDriveByUrl(siteId, folderUrl, logger, args.options.verbose);
      const itemId: string = await drive.getDriveItemId(driveDetails, folderUrl, logger, args.options.verbose);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/drives/${driveDetails.id}/items/${itemId}/createLink`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          type: args.options.type,
          expirationDateTime: args.options.expirationDateTime,
          scope: args.options.scope,
          retainInheritedPermissions: !!args.options.retainInheritedPermissions
        }
      };

      if (args.options.scope === 'users' && args.options.recipients) {
        const recipients = formatting.splitAndTrim(args.options.recipients).map(email => ({ email }));
        requestOptions.data.recipients = recipients;
      }

      const sharingLink = await request.post<any>(requestOptions);

      // remove grantedToIdentities from the sharing link object
      if (sharingLink.grantedToIdentities) {
        delete sharingLink.grantedToIdentities;
      }

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkAddCommand();
