import url from 'url';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import removeCommand, { Options as SpoFileRemoveOptions } from './file-remove.js';

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
    return ['targetUrl', 'sourceUrl'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const webUrl = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    try {
      const serverRelativePath = urlUtil.getServerRelativePath(webUrl, args.options.sourceUrl);
      // Check if the source file exists.
      // Called on purpose, we explicitly check if user specified file
      // in the sourceUrl option.
      // The CreateCopyJobs endpoint accepts file, folder or batch from both.
      // A user might enter folder instead of file as source url by mistake
      // then there are edge cases when deleteIfAlreadyExists flag is set
      // the user can receive misleading error message.
      await this.fileExists(webUrl, serverRelativePath);

      if (args.options.deleteIfAlreadyExists) {
        // try delete target file, if deleteIfAlreadyExists flag is set
        const filename: string = args.options.sourceUrl.replace(/^.*[\\\/]/, '');
        await this.recycleFile(tenantUrl, args.options.targetUrl, filename, logger);
      }

      // all preconditions met, now create copy job
      const sourceAbsoluteUrl: string = urlUtil.urlCombine(tenantUrl, serverRelativePath);
      const allowSchemaMismatch: boolean = args.options.allowSchemaMismatch || false;
      const requestUrl: string = urlUtil.urlCombine(webUrl, '/_api/site/CreateCopyJobs');
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata=nometadata'
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
  private fileExists(webUrl: string, sourceUrl: string): Promise<void> {
    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(sourceUrl)}')/`;
    const requestOptions: CliRequestOptions = {
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
  private async recycleFile(tenantUrl: string, targetUrl: string, filename: string, logger: Logger): Promise<void> {
    const targetFolderAbsoluteUrl: string = urlUtil.urlCombine(tenantUrl, targetUrl);

    // since the target WebFullUrl is unknown we can use getRequestDigest
    // to get it from target folder absolute url.
    // Similar approach used here Microsoft.SharePoint.Client.Web.WebUrlFromFolderUrlDirect
    const contextResponse = await spo.getRequestDigest(targetFolderAbsoluteUrl);

    if (this.debug) {
      await logger.logToStderr(`contextResponse.WebFullUrl: ${contextResponse.WebFullUrl}`);
    }

    const targetFileServerRelativeUrl: string = `${urlUtil.getServerRelativePath(contextResponse.WebFullUrl, targetUrl)}/${filename}`;

    const removeOptions: SpoFileRemoveOptions = {
      webUrl: contextResponse.WebFullUrl,
      url: targetFileServerRelativeUrl,
      recycle: true,
      force: true,
      debug: this.debug,
      verbose: this.verbose
    };

    try {
      await Cli.executeCommand(removeCommand as Command, { options: { ...removeOptions, _: [] } });
    }
    catch (err: any) {
      if (err !== undefined && err.message !== undefined && err.message.includes('does not exist')) {

      }
      else {
        throw err;
      }
    }
  }
}

export default new SpoFileMoveCommand();
