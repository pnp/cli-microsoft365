import { AxiosResponse } from 'axios';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as url from 'url';
import { v4 } from 'uuid';
import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli';
import {
  CommandError,
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  sourceFile: string;
  targetFile: string;
}

class FileConvertPdfCommand extends GraphCommand {
  // Graph's drive item URL of the source file
  private sourceFileGraphUrl?: string;

  public get name(): string {
    return commands.CONVERT_PDF;
  }

  public get description(): string {
    return 'Converts the specified file to PDF using Microsoft Graph';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    let sourceFileUrl: string = '';
    // path to the local file that contains the PDF-converted source file
    let localTargetFilePath: string = args.options.targetFile;
    let sourceIsLocalFile: boolean = true;
    let targetIsLocalFile: boolean = true;
    let error: any;

    const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[auth.defaultResource].accessToken);
    if (typeof isAppOnlyAuth === 'undefined') {
      return cb(new CommandError('Unable to determine authentication type'));
    }

    if (args.options.sourceFile.toLowerCase().startsWith('https://')) {
      sourceIsLocalFile = false;
    }

    if (args.options.targetFile.toLowerCase().startsWith('https://')) {
      localTargetFilePath = path.join(os.tmpdir(), v4());
      targetIsLocalFile = false;

      if (this.debug) {
        logger.logToStderr(`Target set to a URL. Will store the temporary converted file at ${localTargetFilePath}`);
      }
    }

    this
      .getSourceFileUrl(logger, args, isAppOnlyAuth)
      .then((_sourceFileUrl: string): Promise<string> => {
        sourceFileUrl = _sourceFileUrl;
        return this.getGraphFileUrl(logger, sourceFileUrl, this.sourceFileGraphUrl);
      })
      .then((graphFileUrl: string): Promise<any> => this.convertFile(logger, graphFileUrl))
      .then(fileResponse => this.writeFileToDisk(logger, fileResponse, localTargetFilePath))
      .then(_ => this.uploadConvertedFileIfNecessary(logger, targetIsLocalFile, localTargetFilePath, args.options.targetFile))
      .then(_ => this.deleteRemoteSourceFileIfNecessary(logger, sourceIsLocalFile, sourceFileUrl),
        // catch the error from any of the previous promises so that we can
        // clean up resources in case something went wrong
        // if this.deleteRemoteSourceFileIfNecessary fails, it won't be caught
        // here, but rather at the end
        err => error = err
      )
      .then(_ => {
        // if the target was a remote file, delete the local temp file
        if (!targetIsLocalFile) {
          if (this.verbose) {
            logger.logToStderr(`Deleting the temporary PDF file at ${localTargetFilePath}...`);
          }

          try {
            fs.unlinkSync(localTargetFilePath);
          }
          catch (e) {
            return cb(e);
          }
        }
        else {
          if (this.debug) {
            logger.logToStderr(`Target is a local path. Not deleting`);
          }
        }

        if (error) {
          this.handleRejectedODataJsonPromise(error, logger, cb);
        }
        else {
          cb();
        }
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  /**
   * Returns web URL of the file to convert to PDF. If the user specified a URL
   * in command's options, returns the specified URL. If the user specified
   * a local file, it will upload the file and return its web URL. If CLI
   * is authenticated as app-only, uploads the file to the default document
   * library in the root site. If the CLI is authenticated as user, uploads the
   * file to the user's OneDrive for Business
   * @param logger Logger instance
   * @param args Command args
   * @param isAppOnlyAuth True if CLI is authenticated in app-only mode
   * @returns Web URL of the file to upload
   */
  private getSourceFileUrl(logger: Logger, args: CommandArgs, isAppOnlyAuth: boolean): Promise<string> {
    if (args.options.sourceFile.toLowerCase().startsWith('https://')) {
      return Promise.resolve(args.options.sourceFile);
    }

    if (this.verbose) {
      logger.logToStderr('Uploading local file temporarily for conversion...');
    }

    const driveUrl: string = `${this.resource}/v1.0/${isAppOnlyAuth ? 'drive/root' : 'me/drive/root'}`;
    // we need the original file extension because otherwise Graph won't be able
    // to convert the file to PDF
    this.sourceFileGraphUrl = `${driveUrl}:/${v4()}${path.extname(args.options.sourceFile)}`;
    if (this.debug) {
      logger.logToStderr(`Source is a local file. Uploading to ${this.sourceFileGraphUrl}...`);
    }
    return this.uploadFile(args.options.sourceFile, this.sourceFileGraphUrl);
  }

  /**
   * Uploads the specified local file to a document library using Microsoft Graph
   * @param localFilePath Path to the local file to upload
   * @param targetGraphFileUrl Graph drive item URL of the file to upload
   * @returns Absolute URL of the uploaded file
   */
  private uploadFile(localFilePath: string, targetGraphFileUrl: string): Promise<string> {
    const requestOptions: any = {
      url: `${targetGraphFileUrl}:/createUploadSession`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .post<{ uploadUrl: string; expirationDateTime: string }>(requestOptions)
      .then((res: { uploadUrl: string; expirationDateTime: string }): Promise<{ webUrl: string }> => {
        const fileContents = fs.readFileSync(localFilePath);
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
   * @param fileGraphUrl If set, will return this URL without further action
   * @returns Graph's drive item URL for the specified file
   */
  private getGraphFileUrl(logger: Logger, fileWebUrl: string, fileGraphUrl: string | undefined): Promise<string> {
    if (this.debug) {
      logger.logToStderr(`Resolving Graph drive item URL for ${fileWebUrl}`);
    }

    if (fileGraphUrl) {
      if (this.debug) {
        logger.logToStderr(`Returning previously resolved Graph drive item URL ${fileGraphUrl}`);
      }
      return Promise.resolve(fileGraphUrl);
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

  /**
   * Requests conversion of a file to PDF using Microsoft Graph
   * @param logger Logger instance
   * @param graphFileUrl Graph drive item URL of the file to convert to PDF
   * @returns Response object with a URL in the Location header that contains
   * the file converted to PDF. The URL must be called anonymously
   */
  private convertFile(logger: Logger, graphFileUrl: string): Promise<any> {
    if (this.verbose) {
      logger.logToStderr('Converting file...');
    }

    const requestOptions: any = {
      url: `${graphFileUrl}:/content?format=pdf`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'stream'
    };

    return request.get(requestOptions);
  }

  /**
   * Writes the contents of the specified file stream to a local file
   * @param logger Logger instance
   * @param fileResponse Response with stream file contents
   * @param localFilePath Local file path where to store the file
   */
  private writeFileToDisk(logger: Logger, fileResponse: AxiosResponse, localFilePath: string): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Writing converted PDF file to ${localFilePath}...`);
    }

    return new Promise((resolve, reject) => {
      // write the downloaded file to disk
      const writer = fs.createWriteStream(localFilePath);
      fileResponse.data.pipe(writer);

      writer.on('error', err => {
        reject(err);
      });
      writer.on('close', () => {
        resolve();
      });
    });
  }

  /**
   * If the user specified a URL as the targetFile, uploads the converted PDF
   * file to the specified location. If targetFile is a local path, doesn't do
   * anything.
   * @param logger Logger instance
   * @param targetIsLocalFile Boolean that denotes if user specified as the target location a local path
   * @param localFilePath Local file path to where the file to be uploaded is located
   * @param targetFileUrl Web URL of the file to upload
   */
  private uploadConvertedFileIfNecessary(logger: Logger, targetIsLocalFile: boolean, localFilePath: string, targetFileUrl: string): Promise<void> {
    // if the target was a local path, we're done.
    // Otherwise, upload the file to the specified URL
    if (targetIsLocalFile) {
      if (this.debug) {
        logger.logToStderr('Specified target is a local file. Not uploading.');
      }
      return Promise.resolve();
    }

    if (this.verbose) {
      logger.logToStderr(`Uploading converted PDF file to ${targetFileUrl}...`);
    }

    return this
      .getGraphFileUrl(logger, targetFileUrl, undefined)
      .then(targetGraphFileUrl => this.uploadFile(localFilePath, targetGraphFileUrl))
      .then(_ => Promise.resolve());
  }

  /**
   * If the user specified local file to be converted to PDF, removes the file
   * that was temporarily upload to a document library for the conversion.
   * If the specified source file was a URL, doesn't do anything.
   * @param logger Logger instance
   * @param sourceIsLocalFile Boolean that denotes if user specified a local path as the source file
   * @param sourceFileUrl Web URL of the temporary source file to delete
   */
  private deleteRemoteSourceFileIfNecessary(logger: Logger, sourceIsLocalFile: boolean, sourceFileUrl: string): Promise<void> {
    // if the source was a remote file, we're done,
    // otherwise delete the temporary uploaded file
    if (!sourceIsLocalFile) {
      if (this.debug) {
        logger.logToStderr('Source file was URL. Not removing.');
      }
      return Promise.resolve();
    }

    if (this.verbose) {
      logger.logToStderr(`Deleting the temporary file at ${sourceFileUrl}...`);
    }

    return this
      .getGraphFileUrl(logger, sourceFileUrl, this.sourceFileGraphUrl)
      .then((graphFileUrl: string): Promise<void> => {
        const requestOptions: any = {
          url: graphFileUrl,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        return request.delete(requestOptions);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --sourceFile <sourceFile>'
      },
      {
        option: '-t, --targetFile <targetFile>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.sourceFile.toLowerCase().startsWith('https://') &&
      !fs.existsSync(args.options.sourceFile)) {
      // assume local path
      return `Specified source file ${args.options.sourceFile} doesn't exist`;
    }

    if (!args.options.targetFile.toLowerCase().startsWith('https://') &&
      fs.existsSync(args.options.targetFile)) {
      // assume local path
      return `Another file found at ${args.options.targetFile}`;
    }

    return true;
  }
}

module.exports = new FileConvertPdfCommand();