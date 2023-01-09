import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
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
  type: string;
  expirationDateTime?: string;
  scope?: string;
}

class SpoFileSharingLinkAddCommand extends SpoCommand {
  private static readonly types: string[] = ['view', 'edit'];
  private static readonly scopes: string[] = ['anonymous', 'organization'];

  public get name(): string {
    return commands.FILE_SHARINGLINK_ADD;
  }

  public get description(): string {
    return 'Creates a new sharing link for a file';
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
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--fileId [fileId]'
      },
      {
        option: '--fileUrl [fileUrl]'
      },
      {
        option: '--type <type>',
        autocomplete: SpoFileSharingLinkAddCommand.types
      },
      {
        option: '--expirationDateTime [expirationDateTime]'
      },
      {
        option: '--scope [scope]',
        autocomplete: SpoFileSharingLinkAddCommand.scopes
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

        if (SpoFileSharingLinkAddCommand.types.indexOf(args.options.type) < 0) {
          return `'${args.options.type}' is not a valid type. Allowed types are ${SpoFileSharingLinkAddCommand.types.join(', ')}`;
        }

        if (args.options.scope &&
          SpoFileSharingLinkAddCommand.scopes.indexOf(args.options.scope) < 0) {
          return `'${args.options.scope}' is not a valid scope. Allowed scopes are ${SpoFileSharingLinkAddCommand.scopes.join(', ')}`;
        }

        if (args.options.scope && args.options.scope !== 'anonymous' && args.options.expirationDateTime) {
          return `Option expirationDateTime can only be used for links with scope 'anonymous'.`;
        }

        if (args.options.expirationDateTime && !validation.isValidISODateTime(args.options.expirationDateTime)) {
          return `${args.options.expirationDateTime} is not a valid ISO date string.`;
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
      logger.logToStderr(`Creating a sharing link for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await this.getFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);

      const requestOptions: CliRequestOptions = {
        url: `https://graph.microsoft.com/v1.0/sites/${fileDetails.SiteId}/drives/${fileDetails.VroomDriveID}/items/${fileDetails.VroomItemID}/createLink`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          type: args.options.type,
          expirationDateTime: args.options.expirationDateTime,
          scope: args.options.scope
        }
      };

      const sharingLink = await request.post<any>(requestOptions);

      logger.log(sharingLink);
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

module.exports = new SpoFileSharingLinkAddCommand();
