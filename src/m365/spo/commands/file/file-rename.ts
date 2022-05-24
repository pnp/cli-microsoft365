import * as url from 'url';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetFileName: string;
  force: boolean;
}

class SpoFileRenameCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_RENAME;
  }

  public get description(): string {
    return 'Renames a file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.webUrl = typeof args.options.webUrl !== 'undefined';
    telemetryProps.sourceUrl = typeof args.options.sourceUrl !== 'undefined';
    telemetryProps.targetFileName = typeof args.options.targetFileName !== 'undefined';
    telemetryProps.force = typeof args.options.force !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const webUrl = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;
    const webServerRelativeUrl: string = webUrl.replace(tenantUrl, '');
    const originalFileServerRelativeUrl: string = `${webServerRelativeUrl}${args.options.sourceUrl}`;
    
    // Check if the source file exists.
    // Called on purpose, we explicitly check if user specified file
    // in the sourceUrl option.
    this
      .fileExists(originalFileServerRelativeUrl, args.options.webUrl)
      .then((): Promise<void> => {
        const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.sourceUrl.substring(0, args.options.sourceUrl.lastIndexOf('/'))) + '/' + args.options.targetFileName;

        // Check if file already exists with the new name, if so: delete
        if (args.options.force) {
          // delete the target file if force is set
          return this.recycleFile(args.options.webUrl, serverRelativeUrl, logger); 
        }

        return Promise.resolve();
      })
      .then((): Promise<void> => {
        // all preconditions met, now rename item
        const requestBody: any = {
          formValues : [{
            FieldName: 'FileLeafRef',
            FieldValue: args.options.targetFileName
          }]
        };

        const requestOptions: any = {
          url: `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(originalFileServerRelativeUrl)}')/ListItemAllFields/ValidateUpdateListItem()`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: requestBody,
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then((resp: any): Promise<any> => {
        return new Promise<void>((resolve: () => void): void => {
          logger.log(resp.value);
          resolve();
        });
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  /**
   * Checks if the original file exists
   */
  private fileExists(originalFileServerRelativeUrl: string, webUrl: string): Promise<void> {
    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(originalFileServerRelativeUrl)}')`;
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
   * Deletes file in the site recycle bin
   */
  private recycleFile(webUrl: string, targetFileServerRelUrl: string, logger: Logger): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const requestUrl: string = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(targetFileServerRelUrl)}')/Recycle()`;
      const requestOptions: any = {
        url: requestUrl,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request.post(requestOptions)
        .then((): void => {
          resolve();
        })
        .catch((err: any): any => {
          if (err.message === "Request failed with status code 404") {
            // file does not exist so can proceed
            return resolve();
          }

          if (this.debug) {
            logger.logToStderr(`recycleFile error...`);
            logger.logToStderr(err);
          }

          reject(err);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>'
      },
      {
        option: '-t, --targetFileName <targetFileName>'
      },
      {
        option: '--force'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return validation.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoFileRenameCommand();
