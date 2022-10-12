import * as url from 'url';
import Command from '../../../../Command';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Options as SpoFileRemoveOptions } from './file-remove';
const removeCommand: Command = require('./file-remove');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  deleteIfAlreadyExists?: boolean;
  allowSchemaMismatch?: boolean;
}

class SpoFileMoveCommand extends SpoCommand {
  private dots?: string;

  public get name(): string {
    return commands.FILE_MOVE;
  }

  public get description(): string {
    return 'Moves a file to another location';
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
        deleteIfAlreadyExists: args.options.deleteIfAlreadyExists || false,
        allowSchemaMismatch: args.options.allowSchemaMismatch || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>'
      },
      {
        option: '-t, --targetUrl <targetUrl>'
      },
      {
        option: '--deleteIfAlreadyExists'
      },
      {
        option: '--allowSchemaMismatch'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['targetUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const webUrl = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    try {
      // Check if the source file exists.
      // Called on purpose, we explicitly check if user specified file
      // in the sourceUrl option.
      // The CreateCopyJobs endpoint accepts file, folder or batch from both.
      // A user might enter folder instead of file as source url by mistake
      // then there are edge cases when deleteIfAlreadyExists flag is set
      // the user can receive misleading error message.
      this.fileExists(tenantUrl, webUrl, args.options.sourceUrl);

      if (args.options.deleteIfAlreadyExists) {
        // try delete target file, if deleteIfAlreadyExists flag is set
        const filename: string = args.options.sourceUrl.replace(/^.*[\\\/]/, '');
        await this.recycleFile(tenantUrl, args.options.targetUrl, filename, logger);
      }

      // all preconditions met, now create copy job
      const sourceAbsoluteUrl: string = urlUtil.urlCombine(webUrl, args.options.sourceUrl);
      const allowSchemaMismatch: boolean = args.options.allowSchemaMismatch || false;
      const requestUrl: string = urlUtil.urlCombine(webUrl, '/_api/site/CreateCopyJobs');
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          exportObjectUris: [sourceAbsoluteUrl],
          destinationUri: urlUtil.urlCombine(tenantUrl, args.options.targetUrl),
          options: {
            "AllowSchemaMismatch": allowSchemaMismatch,
            "IgnoreVersionHistory": true,
            "IsMoveMode": true
          }
        },
        responseType: 'json'
      };

      const jobInfo = await request.post<any>(requestOptions);
      this.dots = '';

      const copyJobInfo: any = jobInfo.value[0];
      const progressPollInterval: number = 1800; // 30 * 60; //used previously implemented interval. The API does not provide guidance on what value should be used.

      await new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
        setTimeout(() => {
          spo.waitUntilCopyJobFinished({
            copyJobInfo,
            siteUrl: webUrl,
            pollingInterval: progressPollInterval,
            resolve,
            reject,
            logger,
            dots: this.dots,
            debug: this.debug,
            verbose: this.verbose
          });
        }, progressPollInterval);
      });
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Checks if a file exists on the server relative url
   */
  private fileExists(tenantUrl: string, webUrl: string, sourceUrl: string): Promise<void> {
    const webServerRelativeUrl: string = webUrl.replace(tenantUrl, '');
    const fileServerRelativeUrl: string = `${webServerRelativeUrl}${sourceUrl}`;

    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')/`;
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  /**
   * Moves file in the site recycle bin
   */
  private async recycleFile(tenantUrl: string, targetUrl: string, filename: string, logger: Logger) : Promise<void> {
    const targetFolderAbsoluteUrl: string = urlUtil.urlCombine(tenantUrl, targetUrl);

    // since the target WebFullUrl is unknown we can use getRequestDigest
    // to get it from target folder absolute url.
    // Similar approach used here Microsoft.SharePoint.Client.Web.WebUrlFromFolderUrlDirect
    const contextResponse = await spo.getRequestDigest(targetFolderAbsoluteUrl);

    if (this.debug) {
      logger.logToStderr(`contextResponse.WebFullUrl: ${contextResponse.WebFullUrl}`);
    }

    const targetFileServerRelativeUrl: string = `${urlUtil.getServerRelativePath(contextResponse.WebFullUrl, targetUrl)}/${filename}`;

    const removeOptions: SpoFileRemoveOptions = {
      webUrl: contextResponse.WebFullUrl,
      url: targetFileServerRelativeUrl,
      recycle: true,
      confirm: true,
      debug: this.debug,
      verbose: this.verbose
    };

    try {
      await Cli.executeCommand(removeCommand as Command, { options: { ...removeOptions, _: [] } });
    } 
    catch (err: any) {
      if (err.error !== undefined && err.error.message !== undefined && err.error.message.includes('does not exist')) {
        
      }
      else {
        throw err;
      }
    }
  }
}

module.exports = new SpoFileMoveCommand();
