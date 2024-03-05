import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import GraphCommand from '../../base/GraphCommand.js';
import { setTimeout } from 'timers/promises';
import commands from '../commands.js';
import request, { CliRequestOptions } from '../../../request.js';
import { spo } from '../../../utils/spo.js';
import { urlUtil } from '../../../utils/urlUtil.js';
import { validation } from '../../../utils/validation.js';

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

class FileMoveCommand extends GraphCommand {
  private pollingInterval: number = 10_000;
  private readonly nameConflictBehaviorOptions = ['fail', 'replace', 'rename'];

  public get name(): string {
    return commands.MOVE;
  }

  public get description(): string {
    return 'Moves a file to another location using the Microsoft Graph';
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
      const { webUrl, sourceUrl, targetUrl, nameConflictBehavior, newName, verbose } = args.options;
      const sourcePath: string = this.getAbsoluteUrl(webUrl, sourceUrl);
      const destinationPath: string = this.getAbsoluteUrl(webUrl, targetUrl);

      if (verbose) {
        logger.logToStderr(`Moving file '${sourcePath}' to '${destinationPath}'...`);
      }

      const { siteId, driveId, itemId } = await this.getDriveIdAndItemId(webUrl, sourcePath, sourceUrl, logger, verbose);
      const { targetDriveId, targetItemId } = await this.getTargetDriveAndItemId(webUrl, targetUrl, logger, verbose);
      const requestOptions: CliRequestOptions = this.getRequestOptions(targetDriveId, targetItemId, newName, sourcePath);

      const isSameDrive: boolean = driveId === targetDriveId;
      const apiUrl =
        isSameDrive
          ? `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`
          : `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/copy`;

      const queryParameters: string = nameConflictBehavior && nameConflictBehavior !== 'fail'
        ? `@microsoft.graph.conflictBehavior=${nameConflictBehavior}`
        : '';
      const urlWithQuery = `${apiUrl}${queryParameters ? `?${queryParameters}` : ''}`;

      await this.sendRequest(urlWithQuery, requestOptions, isSameDrive, logger);

      if (!isSameDrive) {
        const itemUrl = `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}`;
        await request.delete({ url: itemUrl, headers: requestOptions.headers });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async sendRequest(url: string, requestOptions: CliRequestOptions, isPatch: boolean, logger: Logger): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Moving file...`);
    }
    requestOptions.url = url;

    const response: any = isPatch ? await request.patch(requestOptions) : await request.post(requestOptions);

    if (isPatch) {
      return response;
    }
    else {
      await this.waitUntilCopyOperationCompleted(response.headers.location, logger);
      return response;
    }
  }

  private getAbsoluteUrl(webUrl: string, url: string): string {
    return url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);
  }

  private async getDriveIdAndItemId(webUrl: string, folderUrl: string, sourceUrl: string, logger: Logger, verbose?: boolean): Promise<{ siteId: string, driveId: string, drive: Drive, itemId: string }> {
    const siteId: string = await spo.getSiteId(webUrl, logger, verbose);
    const drive: Drive = await this.getDocumentLibrary(siteId, new URL(folderUrl), sourceUrl);
    const itemId: string = await this.getStartingFolderId(drive, new URL(folderUrl));
    return { siteId, driveId: drive.id as string, drive, itemId };
  }

  private async getDocumentLibrary(siteId: string, folderUrl: URL, folderUrlFromUser: string): Promise<Drive> {
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

  private async getStartingFolderId(documentLibrary: Drive, folderUrl: URL): Promise<string> {
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

  private async getTargetDriveAndItemId(webUrl: string, targetUrl: string, logger: Logger, verbose?: boolean): Promise<{ targetDriveId: string, targetItemId: string }> {
    const targetSiteUrl: string = urlUtil.getTargetSiteAbsoluteUrl(webUrl, targetUrl);
    const targetFolderUrl: string = this.getAbsoluteUrl(targetSiteUrl, targetUrl);
    const { driveId, itemId } = await this.getDriveIdAndItemId(targetSiteUrl, targetFolderUrl, targetUrl, logger, verbose);
    return { targetDriveId: driveId, targetItemId: itemId };
  }

  private getRequestOptions(targetDriveId: string, targetItemId: string, newName: string | undefined, sourcePath: string): CliRequestOptions {
    const requestOptions: CliRequestOptions = {
      url: '',
      headers: { accept: 'application/json;odata.metadata=none' },
      responseType: 'json',
      fullResponse: true,
      data: { parentReference: { driveId: targetDriveId, id: targetItemId } }
    };

    if (newName) {
      const sourceFileName = sourcePath.substring(sourcePath.lastIndexOf('/') + 1);
      const sourceFileExtension = sourceFileName.includes('.') ? sourceFileName.substring(sourceFileName.lastIndexOf('.')) : '';
      const newNameExtension = newName.includes('.') ? newName.substring(newName.lastIndexOf('.')) : '';
      requestOptions.data.name = newNameExtension ? `${newName.replace(newNameExtension, "")}${sourceFileExtension}` : `${newName}${sourceFileExtension}`;
    }

    return requestOptions;
  }

  private async waitUntilCopyOperationCompleted(monitorUrl: string, logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: monitorUrl,
      responseType: 'json'
    };

    const response: any = await request.get(requestOptions);
    if (response.status === 'completed') {
      if (this.verbose) {
        await logger.logToStderr('Copy operation completed succesfully. Returning...');
      }
      return;
    }
    else if (response.status === 'failed') {
      throw response.error.message;
    }
    else {
      if (this.verbose) {
        await logger.logToStderr(`Still copying. Retrying in ${this.pollingInterval / 1000} seconds...`);
      }
      await setTimeout(this.pollingInterval);
      await this.waitUntilCopyOperationCompleted(monitorUrl, logger);
    }
  }
}

export default new FileMoveCommand();