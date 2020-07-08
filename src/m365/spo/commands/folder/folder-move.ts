import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as url from 'url';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
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

class SpoFolderMoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_MOVE;
  }

  public get description(): string {
    return 'Moves a folder to another location';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl: string = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    const sourceAbsoluteUrl: string = this.urlCombine(webUrl, args.options.sourceUrl);
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
          "IgnoreVersionHistory": true,
          "IsMoveMode": true,
        }
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((jobInfo: any): Promise<void> => {
        const jobProgressOptions: JobProgressOptions = {
          webUrl: webUrl,
          copyJopInfo: jobInfo.value[0],
          progressMaxPollAttempts: 1000, // 1 sec.
          progressPollInterval: 30 * 60, // approx. 30 mins. if interval is 1000
          progressRetryAttempts: 5
        };

        return this.getCopyJobProgress(jobProgressOptions, cmd);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  /**
   * A polling function that awaits the 
   * Azure queued copy job to return JobStatus = 0 meaning it is done with the task.
   */
  private getCopyJobProgress(opts: JobProgressOptions, cmd: CommandInstance):
    Promise<void> {
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

      request
        .post<{ JobState?: number, Logs: string[] }>(requestOptions)
        .then((resp: { JobState?: number, Logs: string[] }): void => {
          retryAttemptsCount = 0; // clear retry on promise success 

          if (this.debug) {
            cmd.log('getCopyJobProgress response...');
            cmd.log(resp);
          }

          if (this.verbose) {
            if (resp.JobState && resp.JobState === 4) {
              cmd.log(`Check #${pollCount}. Copy job in progress... JobState: ${resp.JobState}`);
            }
            else {
              cmd.log(`Check #${pollCount}. JobState: ${resp.JobState}`);
            }
          }

          for (const item of resp.Logs) {
            const log: { Event: string; Message: string } = JSON.parse(item);

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
          }
          else if (pollCount < opts.progressMaxPollAttempts) {
            // if the condition isn't met but the timeout hasn't elapsed, go again
            setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
          }
          else {
            reject(new Error('getCopyJobProgress timed out'));
          }
        },
          (error: any) => {
            retryAttemptsCount++;

            // let's retry x times in row before we give up since
            // this is progress check and even if rejects a promise
            // the actual move process can success.
            if (retryAttemptsCount <= opts.progressRetryAttempts) {
              setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
            }
            else {
              reject(error);
            }
          });
    };

    return new Promise<void>(checkCondition);
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
        description: 'The URL of the site where the folder is located'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the folder to move'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to move the folder'
      },
      {
        option: '--allowSchemaMismatch',
        description: 'Ignores any missing fields in the target and moves folder'
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
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
  
    When you move a folder with documents that have version history,
    all of the versions are being moved.
        
  Examples:
  
    Moves folder from a document library located in one site collection to
    another site collection
      ${commands.FOLDER_MOVE} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/

    Moves folder from a document library to another site in the same site
    collection
      ${commands.FOLDER_MOVE} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/
    
    Moves folder to a document library in another site collection. Will ignore any missing fields in the target destination and move anyway
      ${commands.FOLDER_MOVE} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/ --allowSchemaMismatch

  More information:

    Move items from a SharePoint document library
      https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc
    `);
  }
}

module.exports = new SpoFolderMoveCommand();