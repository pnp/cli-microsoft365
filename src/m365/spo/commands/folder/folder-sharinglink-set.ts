import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { drive } from '../../../../utils/drive.js';
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
  id: string;
  expirationDateTime?: string;
}

class SpoFolderSharingLinkSetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_SHARINGLINK_SET;
  }

  public get description(): string {
    return 'Updates a specific sharing link to a folder';
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
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        folderId: typeof args.options.folderId !== 'undefined',
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '--folderUrl [folderUrl]' },
      { option: '--folderId [folderId]' },
      { option: '-i, --id <id>' },
      { option: '--expirationDateTime [expirationDateTime]' }
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

        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return `${args.options.expirationDateTime} is not a valid ISO date string.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['folderId', 'folderUrl'] });
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderId', 'folderUrl', 'id', 'expirationDateTime');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating sharing link to a folder ${args.options.folderId || args.options.folderUrl}...`);
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
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          expirationDateTime: args.options.expirationDateTime
        }
      };

      const sharingLink = await request.patch<any>(requestOptions);

      await logger.log(sharingLink);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFolderSharingLinkSetCommand();