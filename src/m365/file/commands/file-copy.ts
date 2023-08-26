import { Drive, DriveItem, Site } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../cli/Logger.js';
import GlobalOptions from '../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../request.js';
import { formatting } from '../../../utils/formatting.js';
import { urlUtil } from '../../../utils/urlUtil.js';
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
      const sourcePath: string = this.getAbsoluteUrl(args.options.webUrl, args.options.sourceUrl);
      const destinationPath: string = this.getAbsoluteUrl(args.options.webUrl, args.options.targetUrl);

      if (this.verbose) {
        logger.logToStderr(`Copying file '${sourcePath}' to '${destinationPath}'...`);
      }

      const copyUrl: string = await this.getCopyUrl(args.options, sourcePath, logger);
      const { targetDriveId, targetItemId } = await this.getTargetDriveAndItemId(args.options.webUrl, args.options.targetUrl, logger);

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

      if (args.options.newName) {
        requestOptions.data.name = args.options.newName;
      }

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCopyUrl(options: Options, sourceUrl: string, logger: Logger): Promise<string> {
    const folderUrl: URL = new URL(sourceUrl);
    const siteId: string = await this.getSiteId(options.webUrl, logger);
    const drive: Drive = await this.getDocumentLibrary(siteId, folderUrl, options.sourceUrl, logger);
    const itemId: string = await this.getStartingFolderId(drive, folderUrl, logger);

    const queryParameters: string = options.nameConflictBehavior === 'replace'
      ? '@microsoft.graph.conflictBehavior=replace'
      : options.nameConflictBehavior === 'rename'
        ? '@microsoft.graph.conflictBehavior=rename'
        : '';

    const copyUrl: string = `${this.resource}/v1.0/sites/${siteId}/drives/${drive.id}/items/${itemId}/copy${queryParameters ? '?' + queryParameters : ''}`;

    return copyUrl;
  }

  private async getTargetDriveAndItemId(webUrl: string, targetUrl: string, logger: Logger): Promise<{ targetDriveId: string, targetItemId: string }> {
    const targetSiteUrl: string = this.getTargetSiteUrl(webUrl, targetUrl);
    const targetSiteId: string = await this.getSiteId(targetSiteUrl, logger);
    const targetFolderUrl: URL = new URL(this.getAbsoluteUrl(targetSiteUrl, targetUrl));
    const targetDrive: Drive = await this.getDocumentLibrary(targetSiteId, targetFolderUrl, targetUrl, logger);
    const targetDriveId: string = targetDrive.id as string;
    const targetItemId: string = await this.getStartingFolderId(targetDrive, targetFolderUrl, logger);

    return { targetDriveId, targetItemId };
  }

  private async getSiteId(webUrl: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting site id for URL: ${webUrl}...`);
    }

    const url: URL = new URL(webUrl);
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/sites/${formatting.encodeQueryParameter(url.host)}:${url.pathname}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const site: Site = await request.get<Site>(requestOptions);

    return site.id as string;
  }

  private async getDocumentLibrary(siteId: string, folderUrl: URL, folderUrlFromUser: string, logger: Logger): Promise<Drive> {
    if (this.verbose) {
      logger.logToStderr(`Getting document library...`);
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
      logger.logToStderr(`Getting starting folder id...`);
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

  /**
 * Get the site URL from the target SharePoint URL.
 *
 * @param {string} webUrl - The base web URL.
 * @param {string} url - The target SharePoint URL.
 * @returns {string} - The target site URL.
 * 
 * * Example Scenarios:
 * - webUrl = "https://contoso.sharepoint.com" and targetUrl = "/teams/Important/Shared Documents/temp/123/234",
 *    returns "https://contoso.sharepoint.com/teams/Important".
 * - webUrl = "https://contoso.sharepoint.com" and targetUrl = "https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com/Documents/123",
 *    returns "https://contoso-my.sharepoint.com/personal/john_contoso_onmicrosoft_com".
 * - webUrl = "https://contoso.sharepoint.com/teams/finance" and targetUrl = "/Shared Documents/temp",
 *    returns "https://contoso.sharepoint.com".
 * - webUrl = "https://contoso.sharepoint.com" and targetUrl = "/teams/sales/Shared Documents/temp",
 *    returns "https://contoso.sharepoint.com/teams/sales".
 */
  private getTargetSiteUrl(webUrl: string, url: string): string {
    const fullUrl: string = url.startsWith('https://') ? url : urlUtil.getAbsoluteUrl(webUrl, url);

    // Pattern to match SharePoint URLs
    const urlPattern = /https:\/\/[\w\-]+\.sharepoint\.com\/(teams|sites|personal)\/([\w\-]+)/;

    const match = fullUrl.match(urlPattern);

    if (match) {
      // If a match is found, return the matched URL
      return match[0];
    }
    else {
      // Extract the root URL
      const rootUrl = new URL(fullUrl);
      return rootUrl.origin;
    }
  }
}

export default new FileCopyCommand();