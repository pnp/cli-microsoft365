import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
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
  private readonly allowedScopes: string[] = ['anonymous', 'users', 'organization'];

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
        autocomplete: this.allowedScopes
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

        if (args.options.scope && !this.allowedScopes.some(scope => scope === args.options.scope)) {
          return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
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
      if (this.verbose) {
        await logger.logToStderr(`Clearing sharing links from folder ${args.options.folderId || args.options.folderUrl} for ${args.options.scope ? `${args.options.scope} scope` : 'all scopes'}`);
      }

      try {
        const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId, logger, args.options.verbose);
        const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
        const folderUrl: URL = new URL(absoluteFolderUrl);

        const siteId: string = await spo.getSiteId(args.options.webUrl);
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