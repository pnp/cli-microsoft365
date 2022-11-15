import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
}

class SpoFileVersionListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_LIST;
  }

  public get description(): string {
    return 'Retrieves all versions of a file';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '-u, --fileUrl [fileUrl]'
      },
      {
        option: '-i, --fileId [fileId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['fileUrl', 'fileId']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all versions of file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      let requestUrl = `${args.options.webUrl}/_api/web`;
      if (args.options.fileUrl) {
        const serverRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
        requestUrl += `/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
      }
      else {
        requestUrl += `/GetFileById('${args.options.fileId}')`;
      }
      requestUrl += `/versions`;

      const requestOptions: AxiosRequestConfig = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: any[] }>(requestOptions);
      logger.log(response.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileVersionListCommand();