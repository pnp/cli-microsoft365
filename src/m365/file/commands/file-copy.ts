import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../request.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { spo } from '../../../utils/spo.js';
import { validation } from '../../../utils/validation.js';
import GraphCommand from '../../base/GraphCommand.js';
import commands from '../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  newName?: string;
  nameConflictBehavior?: string;
}

class FileCopyCommand extends GraphCommand {
  private readonly nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];

  public get name(): string {
    return commands.COPY;
  }

  public get description(): string {
    return 'Copies a file to another location using the Microsoft Graph';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        webUrl: typeof args.options.webUrl !== 'undefined',
        sourceUrl: typeof args.options.sourceUrl !== 'undefined',
        targetUrl: typeof args.options.targetUrl !== 'undefined',
        newName: typeof args.options.newName !== 'undefined',
        nameConflictBehavior: typeof args.options.nameConflictBehavior !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '-s, --sourceUrl <sourceUrl>' },
      { option: '-t, --targetUrl <targetUrl>' },
      { option: '--newName [newName]' },
      { option: '--nameConflictBehavior [nameConflictBehavior]', autocomplete: this.nameConflictBehaviorOptions }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.nameConflictBehavior && this.nameConflictBehaviorOptions.indexOf(args.options.nameConflictBehavior) === -1) {
          return `${args.options.nameConflictBehavior} is not a valid nameConflictBehavior value. Allowed values: ${this.nameConflictBehaviorOptions.join(', ')}.`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const { webUrl, sourceUrl, targetUrl, newName, verbose } = args.options;

      const sourcePath: string = this.getAbsoluteUrl(webUrl, sourceUrl);
      const destinationPath: string = this.getAbsoluteUrl(webUrl, targetUrl);

      if (this.verbose) {
        await logger.logToStderr(`Copying file '${sourcePath}' to '${destinationPath}'...`);
      }

      const copyUrl: string = await this.getCopyUrl(args.options, sourcePath, logger);
      const { targetDriveId, targetItemId } = await this.getTargetDriveAndItemId(webUrl, targetUrl, logger, verbose);

      const requestOptions: CliRequestOptions = {
        url: copyUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          parentReference: {
            driveId: targetDriveId,
            id: targetItemId
          }
        }
      };

      if (newName) {
        const sourceFileName = sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
        const sourceFileExtension = sourceFileName.includes('.') ? sourceFileName.substring(sourceFileName.lastIndexOf('.')) : '';
        const newNameExtension = newName.includes('.') ? newName.substring(newName.lastIndexOf('.')) : '';

        requestOptions.data.name = newNameExtension ? `${newName.replace(newNameExtension, "")}${sourceFileExtension}` : `${newName}${sourceFileExtension}`;
      }

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCopyUrl(options: Options, sourcePath: string, logger: Logger): Promise<string> {
    const { webUrl, sourceUrl, verbose, nameConflictBehavior } = options;
    const folderUrl: URL = new URL(sourcePath);
    const siteId: string = await spo.getSiteIdByMSGraph(webUrl, logger, verbose);
    const drive: Drive = await this.getDocumentLibrary(siteId, folderUrl, sourceUrl, logger);
    const itemId: string = await this.getStartingFolderId(drive, folderUrl, logger);

    const queryParameters: string = nameConflictBehavior && nameConflictBehavior !== 'fail'
      ? `@microsoft.graph.conflictBehavior=${nameConflictBehavior}`
      : '';

    const copyUrl: string = `${this.resource}/v1.0/sites/${siteId}/drives/${drive.id}/items/${itemId}/copy${queryParameters ? `?${queryParameters}` : ''}`;

    return copyUrl;
  }

  private async getTargetDriveAndItemId(webUrl: string, targetUrl: string, logger: Logger, verbose?: boolean): Promise<{ targetDriveId: string, targetItemId: string }> {
    const targetSiteUrl: string = urlUtil.getTargetSiteAbsoluteUrl(webUrl, targetUrl);
    const targetSiteId: string = await spo.getSiteIdByMSGraph(targetSiteUrl, logger, verbose);
    const targetFolderUrl: URL = new URL(this.getAbsoluteUrl(targetSiteUrl, targetUrl));
    const targetDrive: Drive = await this.getDocumentLibrary(targetSiteId, targetFolderUrl, targetUrl, logger);
    const targetDriveId: string = targetDrive.id as string;
    const targetItemId: string = await this.getStartingFolderId(targetDrive, targetFolderUrl, logger);

    return { targetDriveId, targetItemId };
  }

  private async getDocumentLibrary(siteId: string, folderUrl: URL, folderUrlFromUser: string, logger: Logger): Promise<Drive> {
    if (this.verbose) {
      await logger.logToStderr(`Getting document library...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const drives = await request.get<{ value: Drive[] }>(requestOptions);
    const lowerCaseFolderUrl: string = folderUrl.href.toLowerCase();

    const drive: Drive | undefined = drives.value
      .sort((a, b) => (b.webUrl as string).localeCompare(a.webUrl as string))
      .find((d: Drive) => {
        const driveUrl: string = (d.webUrl as string).toLowerCase();
        // ensure that the drive url is a prefix of the folder url
        return lowerCaseFolderUrl.startsWith(driveUrl) &&
          (driveUrl.length === lowerCaseFolderUrl.length ||
            lowerCaseFolderUrl[driveUrl.length] === '/');
      });

    if (!drive) {
      throw `Document library '${folderUrlFromUser}' not found`;
    }

    return drive;
  }

  private async getStartingFolderId(documentLibrary: Drive, folderUrl: URL, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr(`Getting starting folder id...`);
    }

    const documentLibraryRelativeFolderUrl: string = folderUrl.href.replace(new RegExp(`${documentLibrary.webUrl}`, 'i'), '').replace(/\/+$/, '');

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/drives/${documentLibrary.id}/root${documentLibraryRelativeFolderUrl ? `:${documentLibraryRelativeFolderUrl}` : ''}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const folder = await request.get<DriveItem>(requestOptions);

    return folder?.id as string;
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
  }
}

export default new FileCopyCommand();