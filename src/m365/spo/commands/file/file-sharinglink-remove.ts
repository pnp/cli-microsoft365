import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { GraphFileDetails } from './GraphFileDetails';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  id: string;
  confirm?: boolean
}

class SpoFileSharingLinkRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific sharing link of a file';
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
        fileId: typeof args.options.fileId !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '--fileId [fileId]'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileUrl', 'fileId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeSharingLink: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.verbose) {
          logger.logToStderr(`Removing sharing link of file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
        }

        const fileDetails = await this.getNeededFileInformation(args);

        const requestOptions: CliRequestOptions = {
          url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeSharingLink();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove sharing link ${args.options.id} of file ${args.options.fileUrl || args.options.fileId}?`
      });

      if (result.continue) {
        await removeSharingLink();
      }
    }
  }

  private async getNeededFileInformation(args: CommandArgs): Promise<GraphFileDetails> {
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    if (args.options.fileUrl) {
      const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')`;
    }
    else {
      requestUrl += `GetFileById('${args.options.fileId}')`;
    }

    requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId';

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const res = await request.get<GraphFileDetails>(requestOptions);
    return res;
  }
}

module.exports = new SpoFileSharingLinkRemoveCommand();