import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import * as fs from 'fs';
import request, { CliRequestOptions } from '../../../../request';
import * as url from 'url';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  teamId?: string;
  fromDateTime?: string;
  toDateTime?: string;
  licenseModel?: string;
  withAttachments?: boolean;
  folderPath: string;
}

class TeamsMessageExportCommand extends GraphCommand {
  private readonly allowedLicenseModels: string[] = ['A', 'B'];

  public get name(): string {
    return commands.MESSAGE_EXPORT;
  }

  public get description(): string {
    return 'Export Microsoft Teams chat messages for a given user, or a team.';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--fromDateTime [fromDateTime]'
      },
      {
        option: '--toDateTime [toDateTime]'
      },
      {
        option: '--licenseModel [licenseModel]',
        autocomplete: this.allowedLicenseModels
      },
      {
        option: '--withAttachments'
      },
      {
        option: '--folderPath <folderPath>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userId} is not a valid userPrincipalName`;
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.fromDateTime && !validation.isValidISODateTime(args.options.fromDateTime)) {
          return `${args.options.fromDateTime} is not a valid ISO DateTime`;
        }

        if (args.options.toDateTime && !validation.isValidISODateTime(args.options.toDateTime)) {
          return `${args.options.toDateTime} is not a valid ISO DateTime`;
        }

        if (args.options.licenseModel && !this.allowedLicenseModels.some(value => value === args.options.licenseModel)) {
          return `${args.options.licenseModel} is not a valid license model. Allowed values are ${this.allowedLicenseModels.join(',')}`;
        }

        if (!fs.existsSync(args.options.folderPath)) {
          return `Path ${args.options.folderPath} does not exist.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['userId', 'userName', 'teamId']
      },
      {
        options: ['folderPath'],
        runsWhen: (args: CommandArgs) => {
          return args.options.withAttachments !== undefined && args.options.withAttachments;
        }
      },
      {
        options: ['withAttachments'],
        runsWhen: (args: CommandArgs) => {
          return args.options.folderPath !== undefined;
        }
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let baseUrl = `${this.resource}/v1.0/`;

    if (args.options.userId || args.options.userName) {
      baseUrl += `users/${args.options.userId || args.options.userName}/chats`;
    }
    else {
      baseUrl += `teams/${args.options.teamId}/channels`;
    }

    let requestUrl = `${baseUrl}/getAllMessages`;

    const filters = this.getFilters(args.options);
    if (filters.length > 0) {
      requestUrl += `&$filter=${filters.join(' and ')}`;
    }

    try {
      const res: any[] = await odata.getAllItems(requestUrl);
      if (args.options.withAttachments) {
        for (const message of res) {
          if (message.attachments && message.attachments.length > 0) {
            for await (const attachment of message.attachments) {
              const _url = url.parse(attachment['contentUrl']);
              let siteUrl = _url.protocol + '//' + _url.host!;
              if (_url.path!.split('/')[1] === 'sites' || _url.path!.split('/')[1] === 'teams' || _url.path!.split('/')[1] === 'personal') {
                siteUrl += '/' + _url.path!.split('/')[1] + '/' + _url.path!.split('/')[2];
              }

              const requestOptions: CliRequestOptions = {
                url: `${siteUrl}/_api/web/getFileByServerRelativePath(decodedUrl='${_url.path!}')/$value`,
                headers: {},
                responseType: 'stream'
              };
              const file = await request.get<any>(requestOptions);
              const filePath = `${args.options.folderPath}\\${_url.path!.split('/').pop()}`;
              // Not possible to use async/await for this promise
              await new Promise<void>(() => {
                const writer = fs.createWriteStream(filePath);
                file.data.pipe(writer);

                writer.on('error', err => {
                  throw err;
                });
                writer.on('close', () => {
                  if (this.verbose) {
                    logger.logToStderr(`File saved to path ${filePath}`);
                  }
                  return;
                });
              });
            }
          }
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getFilters(options: Options): string[] {
    const filters = [];
    if (options.fromDateTime) {
      filters.push(`createdDateTime ge ${options.fromDateTime}`);
    }
    if (options.toDateTime) {
      filters.push(`createdDateTime lt ${options.toDateTime}`);
    }
    return filters;
  }
}

module.exports = new TeamsMessageExportCommand();