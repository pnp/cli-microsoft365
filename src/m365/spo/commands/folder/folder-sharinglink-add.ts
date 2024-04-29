import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { driveUtil } from '../../../../utils/driveUtil.js';
import { validation } from '../../../../utils/validation.js';
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
  private static readonly types: string[] = ['view', 'edit'];
  private static readonly scopes: string[] = ['anonymous', 'organization', 'users'];

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
        autocomplete: SpoFolderSharingLinkAddCommand.types
      },
      {
        option: '--expirationDateTime [expirationDateTime]'
      },
      {
        option: '--scope [scope]',
        autocomplete: SpoFolderSharingLinkAddCommand.scopes
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

        if (SpoFolderSharingLinkAddCommand.types.indexOf(args.options.type) < 0) {
          return `'${args.options.type}' is not a valid type. Allowed types are ${SpoFolderSharingLinkAddCommand.types.join(', ')}`;
        }

        if (args.options.scope &&
          SpoFolderSharingLinkAddCommand.scopes.indexOf(args.options.scope) < 0) {
          return `'${args.options.scope}' is not a valid scope. Allowed scopes are ${SpoFolderSharingLinkAddCommand.scopes.join(', ')}`;
        }

        if (args.options.scope && args.options.scope === 'users' && !args.options.recipients) {
          return `Option recipients is required with scope 'users'.`;
        }

        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }

        if (args.options.recipients && args.options.recipients.split(',').some(email => !validation.isValidUserPrincipalName(email))) {
          return `${args.options.recipients} contains one or more invalid users`;
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
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteId(args.options.webUrl);
      const drive: Drive = await driveUtil.getDriveByUrl(siteId, folderUrl);
      const itemId: string = await driveUtil.getDriveItemId(drive, folderUrl);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${itemId}/createLink`,
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
        const recipients = args.options.recipients.split(',').map(email => ({ email }));
        requestOptions.data.recipients = recipients;
      }

      const sharingLink = await request.post<any>(requestOptions);

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkAddCommand();
