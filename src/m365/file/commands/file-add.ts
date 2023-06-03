import * as fs from 'fs';
import * as path from 'path';
import * as url from 'url';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../request';
import { validation } from '../../../utils/validation';
import GraphCommand from '../../base/GraphCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderUrl: string;
  filePath: string;
  siteUrl?: string;
}

class FileAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADD;
  }

  public get description(): string {
    return 'Uploads file to the specified site';
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
        siteUrl: typeof args.options.siteUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --folderUrl <folderUrl>' },
      { option: '-p, --filePath <filePath>' },
      { option: '--siteUrl [siteUrl]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!fs.existsSync(args.options.filePath)) {
          return `Specified source file ${args.options.sourceFile} doesn't exist`;
        }

        if (args.options.siteUrl) {
          const isValidSiteUrl = validation.isValidSharePointUrl(args.options.siteUrl);
          if (isValidSiteUrl !== true) {
            return isValidSiteUrl;
          }
        }

        return validation.isValidSharePointUrl(args.options.folderUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let folderUrlWithoutTrailingSlash = args.options.folderUrl;
    if (folderUrlWithoutTrailingSlash.endsWith('/')) {
      folderUrlWithoutTrailingSlash = folderUrlWithoutTrailingSlash.substr(0, folderUrlWithoutTrailingSlash.length - 1);
    }

    try {
      const graphFileUrl = await this.getGraphFileUrl(logger, `${folderUrlWithoutTrailingSlash}/${path.basename(args.options.filePath)}`, args.options.siteUrl);
      await this.uploadFile(args.options.filePath, graphFileUrl);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Uploads the specified local file to a document library using Microsoft Graph
   * @param localFilePath Path to the local file to upload
   * @param targetGraphFileUrl Graph drive item URL of the file to upload
   * @returns Absolute URL of the uploaded file
   */
  private async uploadFile(localFilePath: string, targetGraphFileUrl: string): Promise<string> {
    const fileContents = fs.readFileSync(localFilePath);
    const isEmptyFile = fileContents.length === 0;
    const requestOptions: CliRequestOptions = {
      url: isEmptyFile ? `${targetGraphFileUrl}:/content` : `${targetGraphFileUrl}:/createUploadSession`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    if (isEmptyFile) {
      const res = await request.put<{ webUrl: string }>(requestOptions);
      return res.webUrl;
    }
    else {
      const res = await request.post<{ uploadUrl: string; expirationDateTime: string }>(requestOptions);

      const requestOptionsPut: CliRequestOptions = {
        url: res.uploadUrl,
        headers: {
          'x-anonymous': true,
          'accept': 'application/json;odata.metadata=none',
          'Content-Length': fileContents.length,
          'Content-Range': `bytes 0-${fileContents.length - 1}/${fileContents.length}`
        },
        data: fileContents,
        responseType: 'json'
      };

      const resPut = await request.put<{ webUrl: string }>(requestOptionsPut);
      return resPut.webUrl;
    }
  }

  /**
   * Gets Graph's drive item URL for the specified file. If the user specified
   * a local file to convert to PDF, returns the URL resolved while uploading
   * the file
   * 
   * Example:
   * 
   * fileWebUrl:
   * https://contoso.sharepoint.com/sites/Contoso/site/Shared%20Documents/file.docx
   * 
   * returns:
   * https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,9d1b2174-9906-43ec-8c9e-f8589de047af,f60c833e-71ce-4a5a-b90e-2a7fdb718397/drives/b!k6NJ6ubjYEehsullOeFTcuYME3w1S8xHoHziURdWlu-DWrqz1yBLQI7E7_4TN6fL/root:/file.docx
   * 
   * @param logger Logger instance
   * @param fileWebUrl Web URL of the file for which to get drive item URL
   * @param siteUrl URL of the site to which upload the file. Optional. Specify to suppress lookup.
   * @returns Graph's drive item URL for the specified file
   */
  private async getGraphFileUrl(logger: Logger, fileWebUrl: string, siteUrl?: string): Promise<string> {
    if (this.debug) {
      logger.logToStderr(`Resolving Graph drive item URL for ${fileWebUrl}`);
    }

    const _fileWebUrl = url.parse(fileWebUrl);
    const _siteUrl = url.parse(siteUrl || fileWebUrl);
    const isSiteUrl = typeof siteUrl !== 'undefined';
    let siteId: string = '';
    let driveRelativeFileUrl: string = '';
    const siteInfo = await this.getGraphSiteInfoFromFullUrl(_siteUrl.host as string, _siteUrl.path as string, isSiteUrl);

    siteId = siteInfo.id;
    let siteRelativeFileUrl: string = (_fileWebUrl.path as string).replace(siteInfo.serverRelativeUrl, '');
    // normalize site-relative URLs for root site collections and root sites

    if (!siteRelativeFileUrl.startsWith('/')) {
      siteRelativeFileUrl = '/' + siteRelativeFileUrl;
    }

    const siteRelativeFileUrlChunks: string[] = siteRelativeFileUrl.split('/');
    driveRelativeFileUrl = `/${siteRelativeFileUrlChunks.slice(2).join('/')}`;
    // chunk 0 is empty because the URL starts with /

    const driveId = await this.getDriveId(logger, siteId, siteRelativeFileUrlChunks[1]);
    const graphUrl: string = `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/root:${driveRelativeFileUrl}`;

    if (this.debug) {
      logger.logToStderr(`Resolved URL ${graphUrl}`);
    }

    return graphUrl;
  }

  /**
   * Retrieves the Graph ID and server-relative URL of the specified (sub)site.
   * Automatically detects which path chunks correspond to (sub)site.
   * @param hostName SharePoint host name, eg. contoso.sharepoint.com
   * @param urlPath Server-relative file URL, eg. /sites/site/docs/file1.aspx
   * @param isSiteUrl Set to true to indicate that the specified URL is a site URL
   * @returns ID and server-relative URL of the site denoted by urlPath
   */
  private async getGraphSiteInfoFromFullUrl(hostName: string, urlPath: string, isSiteUrl: boolean): Promise<{ id: string, serverRelativeUrl: string }> {
    const siteId: string = '';
    const urlChunks: string[] = urlPath.split('/');
    return await this.getGraphSiteInfo(hostName, urlChunks, isSiteUrl ? urlChunks.length - 1 : 0, siteId);
  }

  /**
   * Retrieves Graph site ID and server-relative URL of the site specified
   * using chunks from the URL path. Method is being called recursively as long
   * as it can successfully retrieve the site. When retrieving site fails, method
   * will return the last resolved site ID. If no site ID has been retrieved
   * (method fails on the first execution), it will call the reject callback.
   * @param hostName SharePoint host name, eg. contoso.sharepoint.com
   * @param urlChunks Array of chunks from server-relative URL, eg. ['sites', 'site', 'subsite', 'docs', 'file1.aspx']
   * @param currentChunk Current chunk that's being tested, eg. sites
   * @param lastSiteId Last correctly resolved Graph site ID
   * @returns Graph site ID and server-relative URL of the site specified through chunks
   */
  private async getGraphSiteInfo(hostName: string, urlChunks: string[], currentChunk: number, lastSiteId: string): Promise<{ id: string, serverRelativeUrl: string }> {
    let currentPath: string = urlChunks.slice(0, currentChunk + 1).join('/');
    if (currentPath.endsWith('/sites') ||
      currentPath.endsWith('/teams') ||
      currentPath.endsWith('/personal')) {
      return await this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, '');
    }

    if (!currentPath.startsWith('/')) {
      currentPath = '/' + currentPath;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/sites/${hostName}:${currentPath}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const getResult = (id: string, serverRelativeUrl: string): { id: string, serverRelativeUrl: string } => {
      return {
        id,
        serverRelativeUrl
      };
    };

    try {
      const res = await request.get<{ id: string }>(requestOptions);

      if (currentChunk === urlChunks.length - 1) {
        return getResult(res.id, currentPath);
      }
      else {
        return await this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, res.id);
      }
    }
    catch (err) {
      if (lastSiteId) {
        let serverRelativeUrl: string = `${urlChunks.slice(0, currentChunk).join('/')}`;
        if (!serverRelativeUrl.startsWith('/')) {
          serverRelativeUrl = '/' + serverRelativeUrl;
        }
        return getResult(lastSiteId, serverRelativeUrl);
      }
      else {
        throw err;
      }
    }
  }

  /**
   * Returns the Graph drive ID of the specified document library
   * @param graphSiteId Graph ID of the site where the document library is located
   * @param siteRelativeListUrl Server-relative URL of the document library, eg. /sites/site/Documents
   * @returns Graph drive ID of the specified document library
   */
  private async getDriveId(logger: Logger, graphSiteId: string, siteRelativeListUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/sites/${graphSiteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string; webUrl: string }[] }>(requestOptions);
    if (this.debug) {
      logger.logToStderr(`Searching for drive with a URL ending with /${siteRelativeListUrl}...`);
    }
    const drive = res.value.find(d => d.webUrl.endsWith(`/${siteRelativeListUrl}`));
    if (!drive) {
      throw 'Drive not found';
    }

    return drive.id;
  }
}

module.exports = new FileAddCommand();