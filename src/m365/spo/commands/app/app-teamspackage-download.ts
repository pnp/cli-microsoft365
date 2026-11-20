import fs from 'fs';
import path from 'path';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appItemId: z.number().optional(),
  appItemUniqueId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  appName: z.string().optional(),
  fileName: z.string()
    .refine(fileName => !fs.existsSync(fileName), {
      error: e => `File ${e.input} already exists`
    })
    .optional(),
  appCatalogUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .optional()
    .alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.appItemUniqueId, opts.appItemId, opts.appName].filter(x => x !== undefined).length === 1, {
        error: 'Specify appItemUniqueId, appItemId or appName but not multiple',
        params: {
          customCode: 'optionSet',
          options: ['appItemUniqueId', 'appItemId', 'appName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.appCatalogUrl = args.options.appCatalogUrl;
      const appInfo: AppInfo = {
        id: args.options.appItemId ?? undefined,
        packageFileName: args.options.fileName ?? undefined
      };
      if (this.debug) {
        await logger.logToStderr(`appInfo: ${JSON.stringify(appInfo)}`);
      }

      await this.ensureAppInfo(logger, args, appInfo);

      if (this.debug) {
        await logger.logToStderr(`ensureAppInfo: ${JSON.stringify(appInfo)}`);
      }

      await this.loadAppCatalogUrl(logger, args);

      const requestOptions: CliRequestOptions = {
        url: `${this.appCatalogUrl}/_api/web/tenantappcatalog/downloadteamssolution(${appInfo.id})/$value`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'stream'
      };

      const file = await request.get<any>(requestOptions);

      // Not possible to use async/await for this promise
      await new Promise<void>((resolve, reject) => {
        const writer = fs.createWriteStream(appInfo.packageFileName as string);

        file.data.pipe(writer);

        writer.on('error', err => {
          return reject(err);
        });
        writer.on('close', async () => {
          const fileName = appInfo.packageFileName as string;
          if (this.verbose) {
            await logger.logToStderr(`Package saved to ${fileName}`);
          }
          return resolve();
        });
      });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async ensureAppInfo(logger: Logger, args: CommandArgs, appInfo: AppInfo): Promise<void> {
    if (appInfo.id && appInfo.packageFileName) {
      return;
    }

    if (args.options.appName && !appInfo.packageFileName) {
      appInfo.packageFileName = this.getPackageNameFromFileName(args.options.appName);
    }

    await this.loadAppCatalogUrl(logger, args);
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
      url += `GetFolderByServerRelativePath(DecodedUrl='${appCatalogListName}')/files('${formatting.encodeQueryParameter(args.options.appName)}')/ListItemAllFields?$select=Id`;
    }

    const requestOptions: CliRequestOptions = {
      url,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const res = await request.get<{ Id?: string; File?: { Name: string; } }>(requestOptions);

    if (args.options.appItemUniqueId) {
      appInfo.id = parseInt(res.Id as string);
      if (!appInfo.packageFileName) {
        appInfo.packageFileName = this.getPackageNameFromFileName((res.File as { Name: string }).Name as string);
      }
      return;
    }

    if (args.options.appItemId) {
      if (!appInfo.packageFileName) {
        appInfo.packageFileName = this.getPackageNameFromFileName((res.File as { Name: string }).Name as string);
      }
      return;
    }

    appInfo.id = parseInt(res.Id as string);
  }

  private getPackageNameFromFileName(fileName: string): string {
    return `${path.basename(fileName, path.extname(fileName))}.zip`;
  }

  private async loadAppCatalogUrl(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.appCatalogUrl) {
      return;
    }

    const spoUrl = await spo.getSpoUrl(logger, this.debug);
    const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);
    this.appCatalogUrl = appCatalogUrl;
  }
}

export default new SpoAppTeamsPackageDownloadCommand();