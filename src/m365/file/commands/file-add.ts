import * as fs from 'fs';
import * as path from 'path';
import * as url from 'url';
import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import GraphCommand from '../../base/GraphCommand';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderUrl: string;
  filePath: string;
}

class FileAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADD;
  }

  public get description(): string {
    return 'Uploads file to the specified site';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    let folderUrlWithoutTrailingSlash = args.options.folderUrl;
    if (folderUrlWithoutTrailingSlash.endsWith('/')) {
      folderUrlWithoutTrailingSlash = folderUrlWithoutTrailingSlash.substr(0, folderUrlWithoutTrailingSlash.length - 1);
    }

    this
      .getGraphFileUrl(logger, `${folderUrlWithoutTrailingSlash}/${path.basename(args.options.filePath)}`)
      .then(graphFileUrl => this.uploadFile(args.options.filePath, graphFileUrl))
      .then(_ => cb(), rawRes => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  /**
   * Uploads the specified local file to a document library using Microsoft Graph
   * @param localFilePath Path to the local file to upload
   * @param targetGraphFileUrl Graph drive item URL of the file to upload
   * @returns Absolute URL of the uploaded file
   */
  private uploadFile(localFilePath: string, targetGraphFileUrl: string): Promise<string> {
    const fileContents = fs.readFileSync(localFilePath);
    const isEmptyFile = fileContents.length === 0;
    const requestOptions: any = {
      url: isEmptyFile ? `${targetGraphFileUrl}:/content` : `${targetGraphFileUrl}:/createUploadSession`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    if (isEmptyFile) {
      return request
        .put<{ webUrl: string }>(requestOptions)
        .then((res: { webUrl: string }) => res.webUrl);
    }
    else {
      return request
        .post<{ uploadUrl: string; expirationDateTime: string }>(requestOptions)
        .then((res: { uploadUrl: string; expirationDateTime: string }): Promise<{ webUrl: string }> => {
          const requestOptions: any = {
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

          return request.put<{ webUrl: string }>(requestOptions);
        })
        .then((res: { webUrl: string }) => res.webUrl);
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
   * @returns Graph's drive item URL for the specified file
   */
  private getGraphFileUrl(logger: Logger, fileWebUrl: string): Promise<string> {
    if (this.debug) {
      logger.logToStderr(`Resolving Graph drive item URL for ${fileWebUrl}`);
    }

    const _url = url.parse(fileWebUrl);
    let siteId: string = '';
    let driveRelativeFileUrl: string = '';
    return this
      .getGraphSiteInfoFromFullUrl(_url.host as string, _url.path as string)
      .then(siteInfo => {
        siteId = siteInfo.id;
        let siteRelativeFileUrl: string = (_url.path as string).replace(siteInfo.serverRelativeUrl, '');
        // normalize site-relative URLs for root site collections and root sites
        if (!siteRelativeFileUrl.startsWith('/')) {
          siteRelativeFileUrl = '/' + siteRelativeFileUrl;
        }
        const siteRelativeFileUrlChunks: string[] = siteRelativeFileUrl.split('/');
        driveRelativeFileUrl = `/${siteRelativeFileUrlChunks.slice(2).join('/')}`;
        // chunk 0 is empty because the URL starts with /
        return this.getDriveId(logger, siteId, siteRelativeFileUrlChunks[1]);
      })
      .then(driveId => {
        const graphUrl: string = `${this.resource}/v1.0/sites/${siteId}/drives/${driveId}/root:${driveRelativeFileUrl}`;
        if (this.debug) {
          logger.logToStderr(`Resolved URL ${graphUrl}`);
        }
        return graphUrl;
      });
  }

  /**
   * Retrieves the Graph ID and server-relative URL of the specified (sub)site.
   * Automatically detects which path chunks correspond to (sub)site.
   * @param hostName SharePoint host name, eg. contoso.sharepoint.com
   * @param urlPath Server-relative file URL, eg. /sites/site/docs/file1.aspx
   * @returns ID and server-relative URL of the site denoted by urlPath
   */
  private getGraphSiteInfoFromFullUrl(hostName: string, urlPath: string): Promise<{ id: string, serverRelativeUrl: string }> {
    const siteId: string = '';
    const urlChunks: string[] = urlPath.split('/');
    return new Promise((resolve: (siteInfo: { id: string; serverRelativeUrl: string }) => void, reject: (err: any) => void): void => {
      this.getGraphSiteInfo(hostName, urlChunks, 0, siteId, resolve, reject);
    });
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
   * @param resolve Callback method to call when resolving site info succeeded
   * @param reject Callback method to call when resolving site info failed
   * @returns Graph site ID and server-relative URL of the site specified through chunks
   */
  private getGraphSiteInfo(hostName: string, urlChunks: string[], currentChunk: number, lastSiteId: string, resolve: (siteInfo: { id: string; serverRelativeUrl: string }) => void, reject: (err: any) => void): void {
    let currentPath: string = urlChunks.slice(0, currentChunk + 1).join('/');
    if (currentPath.endsWith('/sites') ||
      currentPath.endsWith('/teams') ||
      currentPath.endsWith('/personal')) {
      return this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, '', resolve, reject);
    }
    if (!currentPath.startsWith('/')) {
      currentPath = '/' + currentPath;
    }
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${hostName}:${currentPath}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get<{ id: string }>(requestOptions)
      .then((res: { id: string }) => {
        this.getGraphSiteInfo(hostName, urlChunks, ++currentChunk, res.id, resolve, reject);
      }, err => {
        if (lastSiteId) {
          let serverRelativeUrl: string = `${urlChunks.slice(0, currentChunk).join('/')}`;
          if (!serverRelativeUrl.startsWith('/')) {
            serverRelativeUrl = '/' + serverRelativeUrl;
          }
          resolve({
            id: lastSiteId,
            serverRelativeUrl: serverRelativeUrl
          });
        }
        else {
          reject(err);
        }
      });
  }

  /**
   * Returns the Graph drive ID of the specified document library
   * @param graphSiteId Graph ID of the site where the document library is located
   * @param siteRelativeListUrl Server-relative URL of the document library, eg. /sites/site/Documents
   * @returns Graph drive ID of the specified document library
   */
  private getDriveId(logger: Logger, graphSiteId: string, siteRelativeListUrl: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${graphSiteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { id: string; webUrl: string }[] }>(requestOptions)
      .then((res: { value: { id: string; webUrl: string }[] }) => {
        if (this.debug) {
          logger.logToStderr(`Searching for drive with a URL ending with /${siteRelativeListUrl}...`);
        }
        const drive = res.value.find(d => d.webUrl.endsWith(`/${siteRelativeListUrl}`));
        if (!drive) {
          return Promise.reject('Drive not found');
        }

        return Promise.resolve(drive.id);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-u, --folderUrl <folderUrl>' },
      { option: '-p, --filePath <filePath>' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!fs.existsSync(args.options.filePath)) {
      return `Specified source file ${args.options.sourceFile} doesn't exist`;
    }

    return SpoCommand.isValidSharePointUrl(args.options.folderUrl);
  }
}

module.exports = new FileAddCommand();