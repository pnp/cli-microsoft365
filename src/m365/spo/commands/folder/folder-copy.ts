import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as url from 'url';
import { CommandInstance } from '../../../../cli';

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
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.allowSchemaMismatch = args.options.allowSchemaMismatch || false;
    return telemetryProps;
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
          "IgnoreVersionHistory": true
        }
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((jobInfo: any): Promise<void> => {
        return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          this.dots = '';

          const copyJobInfo: any = jobInfo.value[0];
          const progressPollInterval: number = 30 * 60; //used previously implemented interval. The API does not provide guidance on what value should be used.

          this.timeout = setTimeout(() => {
            this.waitUntilCopyJobFinished(copyJobInfo, webUrl, progressPollInterval, resolve, reject, cmd, this.dots, this.timeout)
          }, progressPollInterval);
        });
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
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
      },
      {
        option: '--allowSchemaMismatch',
        description: 'Ignores any missing fields in the target document library and copies the folder anyway'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoFolderCopyCommand();
