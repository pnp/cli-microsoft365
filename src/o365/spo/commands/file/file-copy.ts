import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { ContextInfo } from '../../spo';
import * as url from 'url';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  deleteIfAlreadyExists?: boolean;
  allowSchemaMismatch: boolean;
}

interface JobProgressOptions {
  webUrl: string;
  /**
   * Response object retrieved from /_api/site/CreateCopyJobs
   */
  copyJopInfo: any;
  /**
   * Poll interval to call /_api/site/GetCopyJobProgress
   */
  progressPollInterval: number;
  /**
   * Max poll intervals to call /_api/site/GetCopyJobProgress
   * after which to give up
   */
  progressMaxPollAttempts: number;
  /**
   * Retry attempts before give up.
   * Give up if /_api/site/GetCopyJobProgress returns 
   * X reject promises in a row
   */
  progressRetryAttempts: number;
}

class SpoFileCopyCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_COPY;
  }

  public get description(): string {
    return 'Copies a file to another location';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.deleteIfAlreadyExists = args.options.deleteIfAlreadyExists || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    // Check if the source file exists.
    // Called on purpose, we explicitly check if user specified file
    // in the sourceUrl option. 
    // The CreateCopyJobs endpoint accepts file, folder or batch from both.
    // A user might enter folder instead of file as source url by mistake
    // then there are edge cases when deleteIfAlreadyExists flag is set
    // the user can receive misleading error message.
    this
      .fileExists(tenantUrl, webUrl, args.options.sourceUrl)
      .then((): Promise<void> => {
        if (args.options.deleteIfAlreadyExists) {
          // try delete target file, if deleteIfAlreadyExists flag is set
          const filename = args.options.sourceUrl.replace(/^.*[\\\/]/, '');
          return this.recycleFile(tenantUrl, args.options.targetUrl, filename, cmd);
        }

        return Promise.resolve();
      })
      .then((): Promise<any> => {
        // all preconditions met, now create copy job
        const sourceAbsoluteUrl = this.urlCombine(webUrl, args.options.sourceUrl);
        const allowSchemaMismatch: boolean = args.options.allowSchemaMismatch || false;
        const requestUrl: string = this.urlCombine(webUrl, '/_api/site/CreateCopyJobs');
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          body: {
            exportObjectUris: [sourceAbsoluteUrl],
            destinationUri: this.urlCombine(tenantUrl, args.options.targetUrl),
            options: {
              "AllowSchemaMismatch": allowSchemaMismatch,
              "IgnoreVersionHistory": true
            }
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((jobInfo: any): Promise<any> => {
        const jobProgressOptions: JobProgressOptions = {
          webUrl: webUrl,
          copyJopInfo: jobInfo.value[0],
          progressMaxPollAttempts: 1000, // 1 sec.
          progressPollInterval: 30 * 60, // approx. 30 mins. if interval is 1000
          progressRetryAttempts: 5
        }

        return this.waitForJobResult(jobProgressOptions, cmd);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
      json: true
    };

    return request.get(requestOptions);
  }

  /**
   * A polling function that awaits the 
   * queued copy job to return JobStatus = 0 meaning it is done with the task.
   */
  private waitForJobResult(opts: JobProgressOptions, cmd: CommandInstance): Promise<void> {
    let pollCount: number = 0;
    let retryAttemptsCount: number = 0;

    const checkCondition = (resolve: () => void, reject: (error: any) => void): void => {
      pollCount++;
      const requestUrl: string = `${opts.webUrl}/_api/site/GetCopyJobProgress`;
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        body: { "copyJobInfo": opts.copyJopInfo },
        json: true
      };

      request.post(requestOptions).then((resp: any): void => {
        retryAttemptsCount = 0; // clear retry on promise success 

        if (this.verbose) {
          if (resp.JobState && resp.JobState === 4) {
            cmd.log(`Check #${pollCount}. Copy job in progress... JobState: ${resp.JobState}`);
          }
          else {
            cmd.log(`Check #${pollCount}. JobState: ${resp.JobState}`);
          }
        }

        for (const item of resp.Logs) {
          const log = JSON.parse(item);

          // reject if progress error 
          if (log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // three possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        // max poll attempts flag raised = reject promise returned
        if (resp.JobState === 0) {
          // job done
          resolve();
          return;
        }

        if (pollCount < opts.progressMaxPollAttempts) {
          // if the condition isn't met but the timeout hasn't elapsed, go again
          setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
        }
        else {
          reject(new Error('waitForJobResult timed out'));
        }
      },
        (error: any): void => {
          retryAttemptsCount++;

          // let's retry x times in row before we give up since
          // this is progress check and even if rejects a promise
          // the actual copy process can success.
          if (retryAttemptsCount <= opts.progressRetryAttempts) {
            setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
          } else {
            reject(error);
          }
        });
    }

    return new Promise<void>(checkCondition);
  }

  /**
   * Moves file in the site recycle bin
   */
  private recycleFile(tenantUrl: string, targetUrl: string, filename: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const targetFolderAbsoluteUrl: string = this.urlCombine(tenantUrl, targetUrl);

      // since the target WebFullUrl is unknown we can use getRequestDigestForSite
      // to get it from target folder absolute url.
      // Similar approach used here Microsoft.SharePoint.Client.Web.WebUrlFromFolderUrlDirect
      this
        .getRequestDigest(targetFolderAbsoluteUrl)
        .then((contextResponse: ContextInfo): void => {
          if (this.debug) {
            cmd.log(`contextResponse.WebFullUrl: ${contextResponse.WebFullUrl}`);
          }

          if (targetUrl.charAt(0) !== '/') {
            targetUrl = `/${targetUrl}`;
          }
          if (targetUrl.lastIndexOf('/') !== targetUrl.length - 1) {
            targetUrl = `${targetUrl}/`;
          }

          const requestUrl: string = `${contextResponse.WebFullUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(`${targetUrl}${filename}`)}')/recycle()`;
          const requestOptions: any = {
            url: requestUrl,
            method: 'POST',
            headers: {
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            },
            json: true
          };

          request.post(requestOptions)
            .then((): void => {
              resolve();
            })
            .catch((err: any): any => {
              if (err.statusCode === 404) {
                // file does not exist so can proceed
                return resolve();
              }

              if (this.debug) {
                cmd.log(`recycleFile error...`);
                cmd.log(err);
              }

              reject(err);
            });
        }, (e: any) => reject(e));
    });
  }

  /**
   * Combines base and relative url considering any missing slashes
   * @param baseUrl https://contoso.com
   * @param relativeUrl sites/abc
   */
  private urlCombine(baseUrl: string, relativeUrl: string): string {
    // remove last '/' of base if exists
    if (baseUrl.lastIndexOf('/') === baseUrl.length - 1) {
      baseUrl = baseUrl.substring(0, baseUrl.length - 1);
    }

    // remove '/' at 0
    if (relativeUrl.charAt(0) === '/') {
      relativeUrl = relativeUrl.substring(1, relativeUrl.length);
    }

    // remove last '/' of next if exists
    if (relativeUrl.lastIndexOf('/') === relativeUrl.length - 1) {
      relativeUrl = relativeUrl.substring(0, relativeUrl.length - 1);
    }

    return `${baseUrl}/${relativeUrl}`;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the file is located'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the file to copy'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to copy the file'
      },
      {
        option: '--deleteIfAlreadyExists',
        description: 'If a file already exists at the targetUrl, it will be moved to the recycle bin. If omitted, the copy operation will be canceled if the file already exists at the targetUrl location'
      },
      {
        option: '--allowSchemaMismatch',
        description: 'Ignores any missing fields in the target document library and copies the file anyway'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.sourceUrl) {
        return 'Required parameter sourceUrl missing';
      }

      if (!args.options.targetUrl) {
        return 'Required parameter targetUrl missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    When you copy a file using the ${chalk.grey(this.name)} command,
    only the latest version of the file is copied.
        
  Examples:
  
    Copy file to a document library in another site collection
      ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/

    Copy file to a document library in the same site collection
      ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test1/HRDocuments/

    Copy file to a document library in another site collection. If a file with
    the same name already exists in the target document library, move it
    to the recycle bin
      ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --deleteIfAlreadyExists
  
    Copy file to a document library in another site collection. Will ignore
    any missing fields in the target destination and copy anyway
      ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --allowSchemaMismatch

  More information:

    Copy items from a SharePoint document library
      https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc
    `);
  }
}

module.exports = new SpoFileCopyCommand();