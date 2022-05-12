import { Drive, DriveItem, Site } from '@microsoft/microsoft-graph-types';
import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../cli';
import {
  CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import { odata, validation } from '../../../utils';
import GraphCommand from '../../base/GraphCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  recursive?: boolean;
}

class FileListCommand extends GraphCommand {
  foldersToGetFilesFrom: string[] = [];

  public get name(): string {
    return commands.LIST;
  }

  public get description(): string {
    return 'Retrieves files from the specified folder and site';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'lastModifiedByUser'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (error?: any) => void): void {
    let webUrl: string = args.options.webUrl;
    if (!webUrl.endsWith('/')) {
      webUrl += '/';
    }
    const folderUrl: URL = new URL(args.options.folderUrl, webUrl);
    let driveId: string = '';

    this
      .getSiteId(args.options.webUrl, logger)
      .then((siteId: string): Promise<Drive> => this.getDocumentLibrary(siteId, folderUrl, args.options.folderUrl, logger))
      .then((drive: Drive): Promise<string> => {
        driveId = drive.id as string;
        return this.getStartingFolderId(drive, folderUrl, logger);
      })
      .then((folderId: string) => {
        if (this.verbose) {
          logger.logToStderr(`Loading folders to get files from...`);
        }

        // add the starting folder to the list of folders to get files from
        this.foldersToGetFilesFrom.push(folderId);

        return this.loadFoldersToGetFilesFrom(folderId, driveId, args.options.recursive);
      })
      .then(_ => {
        if (this.debug) {
          logger.logToStderr(`Folders to get files from: ${this.foldersToGetFilesFrom.join(', ')}`);
        }

        return this.loadFilesFromFolders(driveId, this.foldersToGetFilesFrom, logger);
      })
      .then(files => {
        logger.log(files);
        cb();
      }, err => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getSiteId(webUrl: string, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting site id...`);
    }

    const url: URL = new URL(webUrl);
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/sites/${encodeURIComponent(url.host)}:${url.pathname}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<Site>(requestOptions)
      .then((site: Site): string => {
        if (this.verbose) {
          logger.logToStderr(`Site id: ${site.id}`);
        }

        return site.id as string;
      });
  }

  private getDocumentLibrary(siteId: string, folderUrl: URL, folderUrlFromUser: string, logger: Logger): Promise<Drive> {
    if (this.verbose) {
      logger.logToStderr(`Getting document library...`);
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/sites/${siteId}/drives?$select=webUrl,id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<{ value: Drive[] }>(requestOptions)
      .then((drives: { value: Drive[] }): Promise<Drive> => {
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
          return Promise.reject(`Document library '${folderUrlFromUser}' not found`);
        }

        if (this.verbose) {
          logger.logToStderr(`Document library: ${drive.webUrl}, ${drive.id}`);
        }

        return Promise.resolve(drive);
      });
  }

  private getStartingFolderId(documentLibrary: Drive, folderUrl: URL, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting starting folder id...`);
    }

    const documentLibraryRelativeFolderUrl: string = folderUrl.href.replace(new RegExp(documentLibrary.webUrl as string, 'i'), '');
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/drives/${documentLibrary.id}/root${documentLibraryRelativeFolderUrl.length > 0 ? `:${documentLibraryRelativeFolderUrl}` : ''}?$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<DriveItem>(requestOptions)
      .then((folder: DriveItem): string => {
        if (this.verbose) {
          logger.logToStderr(`Starting folder id: ${folder.id}`);
        }

        return folder.id as string;
      });
  }

  private loadFoldersToGetFilesFrom(folderId: string, driveId: string, recursive: boolean | undefined): Promise<void> {
    if (!recursive) {
      return Promise.resolve();
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/drives/${driveId}/items('${folderId}')/children?$filter=folder ne null&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };
    return request
      .get<{ value: DriveItem[] }>(requestOptions)
      .then((subfolders: { value: DriveItem[] }): Promise<void> => {
        const subfolderIds: string[] = subfolders.value.map((subfolder: DriveItem) => subfolder.id as string);
        this.foldersToGetFilesFrom = this.foldersToGetFilesFrom.concat(subfolderIds);
        return Promise
          .all(subfolderIds.map((subfolderId: string) => this.loadFoldersToGetFilesFrom(subfolderId, driveId, recursive)))
          .then(_ => Promise.resolve());
      });
  }

  private loadFilesFromFolders(driveId: string, folderIds: string[], logger: Logger): Promise<DriveItem[]> {
    if (this.verbose) {
      logger.logToStderr(`Loading files from folders...`);
    }

    let files: DriveItem[] = [];

    return Promise
      .all(folderIds.map((folderId: string): Promise<DriveItem[]> =>
        // get items from folder. Because we can't filter out folders here
        // we need to get all items from the folder and filter them out later
        odata.getAllItems<DriveItem>(`${this.resource}/v1.0/drives/${driveId}/items/${folderId}/children`)))
      .then(res => {
        // flatten data from all promises
        files = files.concat(...res);

        // remove folders from the list of files
        files = files.filter((item: DriveItem) => item.file);
        files.forEach(file => (file as any).lastModifiedByUser = file.lastModifiedBy?.user?.displayName);
        return files;
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-u, --webUrl <webUrl>' },
      { option: '-f, --folderUrl <folderUrl>' },
      { option: '--recursive' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new FileListCommand();