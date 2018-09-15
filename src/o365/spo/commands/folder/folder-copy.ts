import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import * as url from 'url';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
}

interface JobProgressOptions {
  webUrl: string;
  accessToken: string;
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

class SpoFolderCopyCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    const webUrl: string = args.options.webUrl;
    const tenantUrl: string = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        siteAccessToken = accessToken;

        const sourceAbsoluteUrl: string = this.urlCombine(webUrl, args.options.sourceUrl);
        const requestUrl: string = `${webUrl}/_api/site/CreateCopyJobs`;
        const requestOptions: any = {
          url: requestUrl,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: {
            exportObjectUris: [sourceAbsoluteUrl],
            destinationUri: this.urlCombine(tenantUrl, args.options.targetUrl),
            options: { "IgnoreVersionHistory": true }
          },
          json: true
        };

        if (this.debug) {
          cmd.log('CreateCopyJobs request...');
          cmd.log(requestOptions);
        }

        return request.post(requestOptions);
      })
      .then((jobInfo: any): Promise<any> => {
        if (this.debug) {
          cmd.log('CreateCopyJobs response...');
          cmd.log(jobInfo);
        }

        const jobProgressOptions: JobProgressOptions = {
          webUrl: webUrl,
          accessToken: siteAccessToken,
          copyJopInfo: jobInfo.value[0],
          progressMaxPollAttempts: 1000, // 1 sec.
          progressPollInterval: 30 * 60, // approx. 30 mins. if interval is 1000
          progressRetryAttempts: 5
        }

        return this.getCopyJobProgress(jobProgressOptions, cmd);
      })
      .then((resp: any): void => {
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
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${opts.accessToken}`,
          'accept': 'application/json;odata=nometadata'
        }),
        body: { "copyJobInfo": opts.copyJopInfo },
        json: true
      };

      if (this.debug) {
        cmd.log('getCopyJobProgress request...');
        cmd.log(requestOptions);
      }

      request.post(requestOptions).then((resp: any): void => {
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
          // the actual copy process can success.
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
   * Combines base and relative url 
   * considering any missing slashes
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
        description: 'Site-relative URL of the folder to copy'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to copy the folder'
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
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To copy a folder, you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    When you copy a folder with documents that have version history,
    only the latest document version is copied.
        
  Examples:
  
    Copies folder from a document library located in one site collection to
    another site collection
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/

    Copies folder from a document library to another site in the same site
    collection
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/

  More information:

    Copy items from a SharePoint document library
      https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc
    `);
  }
}

module.exports = new SpoFolderCopyCommand();