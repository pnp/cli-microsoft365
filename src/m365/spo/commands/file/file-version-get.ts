import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  label: string;
  fileUrl?: string;
  fileId?: string;
}

class SpoFileVersionGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_VERSION_GET;
  }

  public get description(): string {
    return 'Get a specific version of a file in a SharePoint Document library';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        fileUrl: args.options.fileUrl,
        fileId: args.options.fileId
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-w, --webUrl <webUrl>'
      },
      {
        option: '--label <label>'
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
        if (args.options.fileId && !validation.isValidGuid(args.options.fileId as string)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['fileUrl', 'fileId']);
  }

  #initTypes(): void {
    this.types.string.push('label');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving version ${args.options.label} of the file ${args.options.fileUrl || args.options.fileId} at site ${args.options.webUrl}...`);
    }

    try {
      const version = await this.getVersion(args);
      logger.log(version.value[0]);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  // Gets files from a folder recursively.
  private async getVersion(args: CommandArgs): Promise<any> {
    let requestUrl;
    if (args.options.fileUrl) {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(args.options.fileUrl)}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/GetFileById('${args.options.fileId}')/versions/?$filter=VersionLabel eq '${args.options.label}'`;
    }
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: any[] }>(requestOptions);
    return response;
  }
}

module.exports = new SpoFileVersionGetCommand();