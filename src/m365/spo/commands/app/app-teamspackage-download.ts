import { AxiosRequestConfig } from 'axios';
import * as fs from 'fs';
import * as path from 'path';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  appItemUniqueId?: string;
  appItemId?: number;
  appName?: string;
  fileName?: string;
}

interface AppInfo {
  // Item ID of the app in the app catalog
  id?: number;
  // File name of where the app package will be downloaded to (.zip)
  packageFileName?: string;
}

class SpoAppTeamsPackageDownloadCommand extends SpoAppBaseCommand {
  private appCatalogUrl?: string;

  public get name(): string {
    return commands.APP_TEAMSPACKAGE_DOWNLOAD;
  }

  public get description(): string {
    return 'Downloads Teams app package for an SPFx solution deployed to tenant app catalog';
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
        appCatalogUrl: typeof args.options.appCatalogUrl !== 'undefined',
        appItemUniqueId: typeof args.options.appItemUniqueId !== 'undefined',
        appItemId: typeof args.options.appItemId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined',
        fileName: typeof args.options.fileName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appItemId [appItemId]' },
      { option: '--appItemUniqueId [appItemUniqueId]' },
      { option: '--appName [appName]' },
      { option: '--fileName [fileName]' },
      { option: '-u, --appCatalogUrl [appCatalogUrl]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.appItemUniqueId &&
          !args.options.appItemId &&
          !args.options.appName) {
          return `Specify appItemUniqueId, appItemId or appName`;
        }
    
        if ((args.options.appItemUniqueId && args.options.appItemId) ||
          (args.options.appItemUniqueId && args.options.appName) ||
          (args.options.appItemId && args.options.appName)) {
          return `Specify appItemUniqueId, appItemId or appName but not multiple`;
        }
    
        if (args.options.appItemUniqueId &&
          !validation.isValidGuid(args.options.appItemUniqueId)) {
          return `${args.options.appItemUniqueId} is not a valid GUID`;
        }
    
        if (args.options.appItemId &&
          isNaN(args.options.appItemId)) {
          return `${args.options.appItemId} is not a number`;
        }
    
        if (args.options.fileName &&
          fs.existsSync(args.options.fileName)) {
          return `File ${args.options.fileName} already exists`;
        }
    
        if (args.options.appCatalogUrl) {
          return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this.appCatalogUrl = args.options.appCatalogUrl;
    const appInfo: AppInfo = {
      id: args.options.appItemId ?? undefined,
      packageFileName: args.options.fileName ?? undefined
    };
    if (this.debug) {
      logger.logToStderr(`appInfo: ${JSON.stringify(appInfo)}`);
    }

    this
      .ensureAppInfo(logger, args, appInfo)
      .then(_ => {
        if (this.debug) {
          logger.logToStderr(`ensureAppInfo: ${JSON.stringify(appInfo)}`);
        }

        return this.loadAppCatalogUrl(logger, args);
      })
      .then(_ => {
        const requestOptions: AxiosRequestConfig = {
          url: `${this.appCatalogUrl}/_api/web/tenantappcatalog/downloadteamssolution(${appInfo.id})/$value`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'stream'
        };

        return request.get(requestOptions);
      })
      .then((file: any): Promise<string> => {
        return new Promise((resolve, reject) => {
          const writer = fs.createWriteStream(appInfo.packageFileName as string);

          file.data.pipe(writer);

          writer.on('error', err => {
            reject(err);
          });
          writer.on('close', () => {
            resolve(appInfo.packageFileName as string);
          });
        });
      })
      .then((file: string): void => {
        if (this.verbose) {
          logger.logToStderr(`Package saved to ${file}`);
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private ensureAppInfo(logger: Logger, args: CommandArgs, appInfo: AppInfo): Promise<void> {
    if (appInfo.id && appInfo.packageFileName) {
      return Promise.resolve();
    }

    if (args.options.appName && !appInfo.packageFileName) {
      appInfo.packageFileName = this.getPackageNameFromFileName(args.options.appName);
    }

    return this
      .loadAppCatalogUrl(logger, args)
      .then(_ => {
        const appCatalogListName = 'AppCatalog';
        const serverRelativeAppCatalogListUrl = `${urlUtil.getServerRelativeSiteUrl(this.appCatalogUrl as string)}/${appCatalogListName}`;

        let url: string = `${this.appCatalogUrl}/_api/web/`;
        if (args.options.appItemUniqueId) {
          url += `GetList('${serverRelativeAppCatalogListUrl}')/GetItemByUniqueId('${args.options.appItemUniqueId}')?$expand=File&$select=Id,File/Name`;
        }
        else if (args.options.appItemId) {
          url += `GetList('${serverRelativeAppCatalogListUrl}')/GetItemById(${args.options.appItemId})?$expand=File&$select=File/Name`;
        }
        else if (args.options.appName) {
          url += `getfolderbyserverrelativeurl('${appCatalogListName}')/files('${encodeURIComponent(args.options.appName)}')/ListItemAllFields?$select=Id`;
        }

        const requestOptions: AxiosRequestConfig = {
          url,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get<{ Id?: string; File?: { Name: string; } }>(requestOptions);
      })
      .then(res => {
        if (args.options.appItemUniqueId) {
          appInfo.id = parseInt(res.Id as string);
          if (!appInfo.packageFileName) {
            appInfo.packageFileName = this.getPackageNameFromFileName((res.File as { Name: string }).Name as string);
          }
          return Promise.resolve();
        }

        if (args.options.appItemId) {
          if (!appInfo.packageFileName) {
            appInfo.packageFileName = this.getPackageNameFromFileName((res.File as { Name: string }).Name as string);
          }
          return Promise.resolve();
        }

        // if (args.options.appName)
        // skipped 'if' clause to provide a default code branch
        appInfo.id = parseInt(res.Id as string);
        return Promise.resolve();
      });
  }

  private getPackageNameFromFileName(fileName: string): string {
    return `${path.basename(fileName, path.extname(fileName))}.zip`;
  }

  private loadAppCatalogUrl(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.appCatalogUrl) {
      return Promise.resolve();
    }

    return spo
      .getSpoUrl(logger, this.debug)
      .then(spoUrl => this.getAppCatalogSiteUrl(logger, spoUrl, args))
      .then(appCatalogUrl => {
        this.appCatalogUrl = appCatalogUrl;
      });
  }
}

module.exports = new SpoAppTeamsPackageDownloadCommand();