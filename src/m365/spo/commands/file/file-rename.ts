import { Cli, Logger } from '../../../../cli';
import Command, {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
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
  targetFileName: string;
  force?: boolean;
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
    telemetryProps.force = !!args.options.force;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const webUrl = args.options.webUrl;
    const originalFileServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.sourceUrl);
    
    this
      .fileExists(originalFileServerRelativeUrl, webUrl)
      .then(_ => {
        if (args.options.force) {
          return this.deleteTargetItem(webUrl, args.options.sourceUrl, args.options.targetFileName);
        } 
        return Promise.resolve();
      })
      .then(_ => {
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

  private deleteTargetItem(webUrl: string, sourceUrl: string, targetFileName: string): Promise<void> {
    const targetFileServerRelativeUrl: string = `${urlUtil.getServerRelativePath(webUrl,sourceUrl.substring(0, sourceUrl.lastIndexOf('/')))}/${targetFileName}`;

    const options: SpoFileRemoveOptions = {
      webUrl: webUrl,
      url: targetFileServerRelativeUrl,
      recycle: true,
      confirm: true
    };

    return Cli.executeCommandWithOutput(removeCommand as Command, { options: { ...options, _: [] } })
      .then(_ => {
        return Promise.resolve();
      }, (err: any) => {
        if (err.error !== null && err.error.message !== null && err.error.message.includes('does not exist')) {
          return Promise.resolve();
        }
        return Promise.reject(err);
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
