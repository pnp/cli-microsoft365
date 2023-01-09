import chalk = require('chalk');
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
  expirationDateTime: string;
  scope?: string;
}

class SpoFileSharingLinkAddCommand extends SpoCommand {
  private static type: string[] = ['view', 'edit', 'embed'];
  private static scope: string[] = ['anonymous', 'organization'];

  public get name(): string {
    return commands.FILE_SHARINGLINK_ADD;
  }

  public get description(): string {
    return 'Creates a new sharing link to a file';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'roles', 'link', 'scope'];
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
        option: '-i, --fileId [fileId]'
      },
      {
        option: '-f, --fileUrl [fileUrl]'
      },
      {
        option: '--type <type>]',
        autocomplete: SpoFileSharingLinkAddCommand.type
      },
      {
        option: '--expirationDateTime [expirationDateTime]'
      },
      {
        option: '--scope [scope]',
        autocomplete: SpoFileSharingLinkAddCommand.scope
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

        if (args.options.type &&
          SpoFileSharingLinkAddCommand.type.indexOf(args.options.type) < 0) {
          return `'${args.options.type}' is not a valid type. Allowed types are ${SpoFileSharingLinkAddCommand.type.join(', ')}`;
        }

        if (args.options.scope &&
          SpoFileSharingLinkAddCommand.scope.indexOf(args.options.scope) < 0) {
          return `'${args.options.scope}' is not a valid scope type. Allowed scope types are ${SpoFileSharingLinkAddCommand.scope.join(', ')}`;
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
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Creates a sharing link for file ${args.options.fileId || args.options.fileUrl}...`);
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
          ...(args.options.expirationDateTime && { expirationDateTime: args.options.expirationDateTime }),
          ...(args.options.scope && { scope: args.options.scope })
        }
      };

      const sharingLink: any = await request.post(requestOptions);

      if (!args.options.output || args.options.output === 'json' || args.options.output === 'md') {
        logger.log(sharingLink);
      }
      else {
        //converted to text friendly output
        logger.log({
          id: sharingLink.id,
          roles: sharingLink.roles.join(','),
          link: sharingLink.link.webUrl,
          scope: sharingLink.link.scope
        });
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

module.exports = new SpoFileSharingLinkAddCommand();