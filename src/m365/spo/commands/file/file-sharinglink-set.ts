import chalk = require('chalk');
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  id: string;
  fileId?: string;
  fileUrl?: string;
  role?: string;
  expirationDateTime: string;
}

class SpoFileSharingLinkSetCommand extends SpoCommand {
  private static role: string[] = ['read', 'write'];

  public get name(): string {
    return commands.FILE_SHARINGLINK_SET;
  }

  public get description(): string {
    return 'Updates a sharing link to a file';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'link', 'scope'];
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
        role: typeof args.options.role !== 'undefined'
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
      },
      {
        option: '--id <id>'
      },
      {
        option: '--role [role]',
        autocomplete: SpoFileSharingLinkSetCommand.role
      },
      {
        option: '--expirationDateTime [expirationDateTime]'
      },

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

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.role &&
          SpoFileSharingLinkSetCommand.role.indexOf(args.options.role) < 0) {
          return `'${args.options.role}' is not a valid scope type. Allowed scope types are ${SpoFileSharingLinkSetCommand.role.join(', ')}`;
        }

        const parsedDateTime = Date.parse(args.options.expirationDateTime as string);
        if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
          return `${args.options.expirationDateTime} is not a valid date format. Provide the date in one of the following formats:
      ${chalk.grey('YYYY-MM-DD')}
      ${chalk.grey('YYYY-MM-DDThh:mm')}
      ${chalk.grey('YYYY-MM-DDThh:mmZ')}
      ${chalk.grey('YYYY-MM-DDThh:mm±hh:mm')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['fileId', 'fileUrl'] });
    this.optionSets.push({ options: ['role', 'expirationDateTime'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Updates a sharing link for file ${args.options.fileId || args.options.fileUrl}...`);
    }

    try {
      const fileDetails = await this.getFileDetails(args.options.webUrl, args.options.fileId, args.options.fileUrl);
      const sharingInformation = await this.getSharingInformation(args.options.webUrl, fileDetails, args.options.id);
      if (args.options.role) {
        sharingInformation.role = args.options.role === "read" ? 1 : 2;
      }
      else {
        sharingInformation.expiration = this.parseDateExact(args.options.expirationDateTime);
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/ShareLink?@a1='${fileDetails.ListId}'&@a2='${fileDetails.UniqueId}'`,
        headers: {
          accept: 'application/json;odata=verbose'
        },
        responseType: 'json',
        data: {
          "request": {
            "createLink": true,
            "settings": sharingInformation,
            "emailData": {
              "body": ""
            }
          }
        }
      };

      const sharingLink: any = await request.post(requestOptions);

      if (!args.options.output || args.options.output === 'json' || args.options.output === 'md') {
        logger.log(sharingLink.d.ShareLink.sharingLinkInfo);
      }
      else {
        logger.log({
          id: sharingLink.d.ShareLink.sharingLinkInfo.ShareId,
          link: sharingLink.d.ShareLink.sharingLinkInfo.Url,
          scope: sharingLink.d.ShareLink.sharingLinkInfo.Scope
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
  private async getSharingInformation(webUrl: string, fileDetails: any, id: string): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/Lists(@a1)/GetItemByUniqueId(@a2)/GetSharingInformation?@a1='${fileDetails.ListId}'&@a2='${fileDetails.UniqueId}'&$Expand=sharingLinkTemplates`,
      headers: {
        accept: 'application/json;odata=verbose'
      },
      responseType: 'json'
    };

    const sharingInformation: any = await request.post(requestOptions);
    const sharingInformationOfId: any = sharingInformation.d.sharingLinkTemplates.templates.results.filter((x: any) => x.linkDetails ? x.linkDetails.ShareId === id : false)[0];

    return ({
      expiration: sharingInformationOfId.linkDetails.Expiration,
      linkKind: sharingInformationOfId.linkDetails.LinkKind,
      restrictShareMembership: sharingInformationOfId.linkDetails.RestrictedShareMembership,
      role: sharingInformationOfId.role,
      scope: sharingInformationOfId.scope,
      shareId: id
    });
  }

  private async getFileDetails(webUrl: string, fileId?: string, fileUrl?: string): Promise<any> {
    let requestUrl: string = `${webUrl}/_api/web/`;

    if (fileId) {
      requestUrl += `GetFileById('${fileId}')`;
    }
    else if (fileUrl) {
      requestUrl += `GetFileByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(fileUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl += '?$select=ListId,UniqueId',
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const res = await request.get<any>(requestOptions);
    return res;
  }

  private parseDateExact(date: string): string {
    const d: Date = new Date(Date.parse(date));
    return d.toISOString().split('.')[0].replace(/-/g, '').replace(/:/g, '') + d.toString().split('GMT')[1].split(' ')[0];
  }
}

module.exports = new SpoFileSharingLinkSetCommand();