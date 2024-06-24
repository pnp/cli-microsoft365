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
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl?: string;
  folderId?: string;
  scope?: string;
}

class SpoFolderSharingLinkListCommand extends SpoCommand {
  private readonly allowedScopes: string[] = ['anonymous', 'users', 'organization'];

  public get name(): string {
    return commands.FOLDER_SHARINGLINK_LIST;
  }

  public get description(): string {
    return 'Lists all the sharing links of a specific folder';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'scope', 'roles', 'link'];
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
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        folderId: typeof args.options.folderId !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '--folderUrl [folderUrl]' },
      { option: '--folderId [folderId]' },
      {
        option: '-s, --scope [scope]',
        autocomplete: this.allowedScopes
      }
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

        if (args.options.scope && this.allowedScopes.indexOf(args.options.scope) === -1) {
          return `'${args.options.scope}' is not a valid scope. Allowed values are: ${this.allowedScopes.join(',')}`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'scope');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving sharing links for folder ${args.options.folderId || args.options.folderUrl}...`);
    }

    try {
      const relFolderUrl: string = await spo.getFolderServerRelativeUrl(args.options.webUrl, args.options.folderUrl, args.options.folderId);
      const absoluteFolderUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, relFolderUrl);
      const folderUrl: URL = new URL(absoluteFolderUrl);

      const siteId: string = await spo.getSiteId(args.options.webUrl);
      const drive: Drive = await driveUtil.getDriveByUrl(siteId, folderUrl);
      const itemId: string = await driveUtil.getDriveItemId(drive, folderUrl);

      let requestUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${itemId}/permissions?$filter=Link ne null`;
      if (args.options.scope) {
        requestUrl += ` and Link/Scope eq '${args.options.scope}'`;
      }

      const sharingLinks = await odata.getAllItems<any>(requestUrl);

      if (!args.options.output || !cli.shouldTrimOutput(args.options.output)) {
        await logger.log(sharingLinks);
      }
      else {
        //converted to text friendly output
        await logger.log(sharingLinks.map(i => {
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
