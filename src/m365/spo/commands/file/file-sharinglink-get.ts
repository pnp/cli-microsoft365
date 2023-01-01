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
  id: string;
}

class SpoFileSharinglinkGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_GET;
  }

  public get description(): string {
    return 'Gets details about a specific sharing link';
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
        fileUrl: (!(!args.options.fileUrl)).toString(),
        fileId: (!(!args.options.fileId)).toString()
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

  private async getNeededFileInformation(args: CommandArgs): Promise<{ SiteId: string; VroomDriveID: string; VroomItemID: string; }> {
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    if (args.options.fileId) {
      requestUrl += `GetFileById('${args.options.fileId as string}')`;
    }
    else {
      const fileServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl as string);
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileServerRelativeUrl)}')`;
    }

    requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId';

    const requestOptions: AxiosRequestConfig = {
      url: requestUrl,
      headers: {
        'accept': 'application/json'
      },
      responseType: 'json'
    };

    return await request.get<{ SiteId: string; VroomDriveID: string; VroomItemID: string; }>(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving sharing link for file ${args.options.fileUrl || args.options.fileId} with id ${args.options.id}...`);
    }

    try {
      const fileDetails = await this.getNeededFileInformation(args);

      const requestOptions: AxiosRequestConfig = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions/${args.options.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const res = await request.get(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoFileSharinglinkGetCommand();