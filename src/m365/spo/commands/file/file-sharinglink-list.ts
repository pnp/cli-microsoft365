import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { GraphFileDetails } from './GraphFileDetails';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileId?: string;
  fileUrl?: string;
}

class SpoFileSharingLinkListCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_SHARINGLINK_LIST;
  }

  public get description(): string {
    return 'Lists all the sharing links of a specific file';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'roles', 'link'];
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
        fileId: typeof args.options.fileId !== 'undefined',
        fileUrl: typeof args.options.fileUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --fileId [fileId]'
      },
      {
        option: '-f, --fileUrl [fileUrl]'
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
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving sharing links for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await this.getFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
      const sharingLinks = await odata.getAllItems<any>(`https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/permissions?$filter=Link ne null`);

      if (!args.options.output || args.options.output === 'json' || args.options.output === 'md') {
        logger.log(sharingLinks);
      }
      else {
        //converted to text friendly output
        logger.log(sharingLinks.map(i => {
          return {
            id: i.id,
            roles: i.roles.join(','),
            link: i.link.webUrl
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFileDetails(webUrl: string, fileId?: string, fileUrl?: string): Promise<GraphFileDetails> {
    let requestUrl: string = `${webUrl}/_api/web/`;

    if (fileId) {
      requestUrl += `GetFileById('${fileId}')`;
    }
    else if (fileUrl) {
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl += '?$select=SiteId,VroomItemId,VroomDriveId',
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const res = await request.get<GraphFileDetails>(requestOptions);
    return res;
  }
}

module.exports = new SpoFileSharingLinkListCommand();