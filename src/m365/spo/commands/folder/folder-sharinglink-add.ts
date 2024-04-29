import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl?: string;
  folderId?: string;
  type: string;
  expirationDateTime?: string;
  scope?: string;
  retainInheritiedPermissions?: boolean;
  recipients?: string;
}

class SpoFolderSharingLinkAddCommand extends SpoCommand {
  private readonly allowedTypes: string[] = ['view', 'edit'];
  private readonly allowedScopes: string[] = ['anonymous', 'organization', 'users'];

  public get name(): string {
    return commands.FOLDER_SHARINGLINK_ADD;
  }

  public get description(): string {
    return 'Creates a new sharing link to a folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        folderId: typeof args.options.folderId !== 'undefined',
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        type: typeof args.options.type !== 'undefined',
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        retainInheritiedPermissions: !!args.options.retainInheritiedPermissions,
        recipients: typeof args.options.recipients !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderId [folderId]'
      },
      {
        option: '--folderUrl [folderUrl]'
      },
      {
        option: '--type <type>',
        autocomplete: this.allowedTypes
      },
      {
        option: '--expirationDateTime [expirationDateTime]'
      },
      {
        option: '--scope [scope]',
        autocomplete: this.allowedScopes
      },
      {
        option: '--retainInheritedPermissions [retainInheritedPermissions]'
      },
      {
        option: '--recipients [recipients]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.folderId && !validation.isValidGuid(args.options.folderId)) {
          return `${args.options.folderId} is not a valid GUID`;
        }

        if (args.options.type && !this.allowedTypes.some(type => type === args.options.type)) {
          return `'${args.options.type}' is not a valid type. Allowed values are: ${this.allowedTypes.join(',')}`;
        }

        if (args.options.scope) {
          if (!this.allowedScopes.includes(args.options.scope)) {
            return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(', ')}.`;
          }
          if (args.options.scope === 'users' && !args.options.recipients) {
            return `The 'recipients' option is required when scope is set to 'users'.`;
          }
        }

        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }

        if (args.options.recipients) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.recipients);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'recipients': ${isValidUPNArrayResult}.`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['folderId', 'folderUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderId', 'folderUrl', 'type', 'expirationDateTime', 'scope', 'recipients');
    this.types.boolean.push('retainInheritiedPermissions');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Creating a sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
    }

    try {
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteId(args.options.webUrl);
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
          retainInheritedPermissions: !!args.options.retainInheritiedPermissions
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
