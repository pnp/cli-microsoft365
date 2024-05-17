import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { driveUtil } from '../../../../utils/driveUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { Drive } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../../../../request.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl?: string;
  folderId?: string;
  scope?: string;
  force?: boolean;
}

class SpoFolderSharingLinkClearCommand extends SpoCommand {
  private static readonly scopes: string[] = ['anonymous', 'organization', 'users'];

  public get name(): string {
    return commands.FOLDER_SHARINGLINK_CLEAR;
  }

  public get description(): string {
    return 'Removes all sharing links of a folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        webUrl: typeof args.options.webUrl !== 'undefined',
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        folderId: typeof args.options.folderId !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '--folderUrl [folderUrl]' },
      { option: '--folderId [folderId]' },
      {
        option: '--scope [scope]',
        autocomplete: SpoFolderSharingLinkClearCommand.scopes
      },
      { option: '-f, --force' }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['folderUrl', 'folderId'] }
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

        if (args.options.scope &&
          SpoFolderSharingLinkClearCommand.scopes.indexOf(args.options.scope) < 0) {
          return `'${args.options.scope}' is not a valid scope. Allowed scopes are ${SpoFolderSharingLinkClearCommand.scopes.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'scope');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const clearSharingLinks = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Clearing sharing links from folder ${args.options.folderId || args.options.folderUrl} for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}`);
        }

        const { drive, itemId } = await this.getDriveAndItemId(args.options);
        const sharingLinks = await this.getSharingLinks(drive.id, itemId, args.options.scope);

        const requestOptions: CliRequestOptions = {
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        for (const sharingLink of sharingLinks) {
          requestOptions.url = `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${itemId}/permissions/${sharingLink.id}`;
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

  private async getDriveAndItemId(options: Options): Promise<{ drive: Drive, itemId: string }> {
    const relFolderUrl: string = await spo.getFolderServerRelativeUrl(options.webUrl, options.folderUrl, options.folderId);
    const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(options.webUrl, relFolderUrl);
    const folderUrl: URL = new URL(absoluteFolderUrl);

    const siteId: string = await spo.getSiteId(options.webUrl);
    const drive: Drive = await driveUtil.getDriveByUrl(siteId, folderUrl);
    const itemId: string = await driveUtil.getDriveItemId(drive, folderUrl);
    return { drive, itemId };
  }

  private async getSharingLinks(driveId: string | undefined, itemId: string, scope?: string): Promise<any[]> {
    let requestUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/permissions?$filter=Link ne null`;
    if (scope) {
      requestUrl += ` and Link/Scope eq '${scope}'`;
    }
    return odata.getAllItems<any>(requestUrl);
  }
}

export default new SpoFolderSharingLinkClearCommand();