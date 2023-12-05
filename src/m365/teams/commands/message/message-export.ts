import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import * as fs from 'fs';
import request, { CliRequestOptions } from '../../../../request';
import * as url from 'url';
import { urlUtil } from '../../../../utils/urlUtil';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  teamId?: string;
  fromDateTime?: string;
  toDateTime?: string;
  withAttachments?: boolean;
  folderPath?: string;
}

class TeamsMessageExportCommand extends GraphCommand {

  public get name(): string {
    return commands.MESSAGE_EXPORT;
  }

  public get description(): string {
    return 'Exports Microsoft Teams chat messages for a given user, or a team.';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        teamId: typeof args.options.teamId !== 'undefined',
        fromDateTime: typeof args.options.fromDateTime !== 'undefined',
        toDateTime: typeof args.options.toDateTime !== 'undefined',
        withAttachments: !!args.options.withAttachments,
        folderPath: typeof args.options.folderPath !== 'undefined'
      });
    });
  }


  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --userId [userId]'
      },
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '-f, --fromDateTime [fromDateTime]'
      },
      {
        option: '-t, --toDateTime [toDateTime]'
      },
      {
        option: '-a, --withAttachments'
      },
      {
        option: '-p, --folderPath [folderPath]'
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
          return `${args.options.userName} is not a valid userPrincipalName`;
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

        if (args.options.folderPath && !fs.existsSync(args.options.folderPath)) {
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
          return args.options.withAttachments === true;
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

    let requestUrl = args.options.userId || args.options.userName
      ? `${this.resource}/v1.0/users/${args.options.userId || args.options.userName}/chats/getAllMessages`
      : `${this.resource}/v1.0/teams/${args.options.teamId}/channels/getAllMessages`;

    const filters = this.getFilters(args.options);

    if (filters.length > 0) {
      requestUrl += `?$filter=${filters.join(' and ')}`;
    }

    try {
      const res: any[] = await odata.getAllItems(requestUrl);

      if (args.options.withAttachments) {
        for (const message of res) {

          if (message.attachments?.length > 0) {
            for await (const attachment of message.attachments) {
              const contentUrl = url.parse(attachment['contentUrl']);
              let siteUrl = contentUrl.protocol + '//' + contentUrl.host!;

              if (contentUrl.path!.split('/')[1] === 'sites' || contentUrl.path!.split('/')[1] === 'teams' || contentUrl.path!.split('/')[1] === 'personal') {
                siteUrl += '/' + contentUrl.path!.split('/')[1] + '/' + contentUrl.path!.split('/')[2];
              }

              const serverRelativePath = urlUtil.getServerRelativePath(siteUrl, contentUrl.path!);

              const requestOptions: CliRequestOptions = {
                url: `${siteUrl}/_api/web/getFileByServerRelativePath(decodedUrl=@decodedUrl)/$value?@decodedUrl='${serverRelativePath}'`,
                headers: {},
                responseType: 'stream'
              };
              const file = await request.get<any>(requestOptions);
              const folderPath = `${args.options.folderPath}\\${message.id}`;
              const filePath = `${folderPath}\\${contentUrl.path!.split('/').pop()}`;

              if (!fs.existsSync(folderPath)) {
                fs.mkdirSync(folderPath);
              }

              await new Promise<void>((resolve, reject) => {
                const writer = fs.createWriteStream(filePath);
                file.data.pipe(writer);

                writer.on('error', err => {
                  reject(err);
                });

                writer.on('close', () => {
                  if (this.verbose) {
                    logger.logToStderr(`File saved to path ${filePath}`);
                  }
                  resolve();
                });
              });
            }
          }
        }
      }
      logger.log(res);
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