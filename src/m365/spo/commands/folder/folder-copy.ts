import url from 'url';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  allowSchemaMismatch?: boolean;
}

class SpoFolderCopyCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
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
    const webUrl: string = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    const serverRelativePath = urlUtil.getServerRelativePath(webUrl, args.options.sourceUrl);
    const sourceAbsoluteUrl: string = urlUtil.urlCombine(tenantUrl, serverRelativePath);
    const allowSchemaMismatch: boolean = args.options.allowSchemaMismatch || false;
    const requestOptions: CliRequestOptions = {
      url: urlUtil.urlCombine(webUrl, '/_api/site/CreateCopyJobs'),
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: {
        exportObjectUris: [sourceAbsoluteUrl],
        destinationUri: urlUtil.urlCombine(tenantUrl, args.options.targetUrl),
        options: {
          "AllowSchemaMismatch": allowSchemaMismatch,
          "IgnoreVersionHistory": true
        }
      },
      responseType: 'json'
    };

    try {
      const jobInfo = await request.post<any>(requestOptions);
      const copyJobInfo: any = jobInfo.value[0];
      const progressPollInterval: number = 30 * 60; //used previously implemented interval. The API does not provide guidance on what value should be used.


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
}

export default new SpoFolderCopyCommand();
