import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { spo } from '../../../../utils/spo.js';
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
  id: string;
  force?: boolean;
}

class SpoFolderSharingLinkRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_REMOVE;
  }

  public get description(): string {
    return 'Removes a sharing link from a folder';
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
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '--folderUrl [folderUrl]' },
      { option: '--folderId [folderId]' },
      { option: '-i, --id <id>' },
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

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl', 'folderId', 'id');
    this.types.boolean.push('force');
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
