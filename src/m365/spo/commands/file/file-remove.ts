import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id?: string;
  url?: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoFileRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.url = (!(!args.options.url)).toString();
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return ['url'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeFile: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing file in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.id) {
        requestUrl = `${args.options.webUrl}/_api/web/GetFileById(guid'${encodeURIComponent(args.options.id as string)}')`;
      }
      else {
        // concatenate trailing '/' if not provided
        // so if the provided url is for the root site, the substr bellow will get the right value
        let serverRelativeSiteUrl: string = args.options.webUrl;
        if (!serverRelativeSiteUrl.endsWith('/')) {
          serverRelativeSiteUrl = `${serverRelativeSiteUrl}/`;
        }
        serverRelativeSiteUrl = serverRelativeSiteUrl.substr(serverRelativeSiteUrl.indexOf('/', 8));

        let fileUrl: string = args.options.url as string;
        if (!fileUrl.startsWith(serverRelativeSiteUrl)) {
          fileUrl = `${serverRelativeSiteUrl}${fileUrl}`;
        }
        requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`;
      }

      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }

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

      request
        .post(requestOptions)
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeFile();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the file ${args.options.id || args.options.url} located in site ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeFile();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-u, --url [url]'
      },
      {
        option: '--recycle'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id &&
      !validation.isValidGuid(args.options.id as string)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.id && args.options.url) {
      return 'Specify id or url, but not both';
    }

    if (!args.options.id && !args.options.url) {
      return 'Specify id or url';
    }

    return true;
  }
}

module.exports = new SpoFileRemoveCommand();