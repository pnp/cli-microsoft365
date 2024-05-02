import { Message } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Outlook } from '../../Outlook.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  folderId?: string;
  folderName?: string;
  startTime?: string;
  endTime?: string;
  userId?: string;
  userName?: string;
}

class OutlookMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_LIST;
  }

  public get description(): string {
    return 'Gets all mail messages from the specified folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        folderId: typeof args.options.folderId !== 'undefined',
        folderName: typeof args.options.folderName !== 'undefined',
        startTime: typeof args.options.startTime !== 'undefined',
        endTime: typeof args.options.endTime !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--folderName [folderName]',
        autocomplete: Outlook.wellKnownFolderNames
      },
      {
        option: '--folderId [folderId]'
      },
      {
        option: '--startTime [startTime]'
      },
      {
        option: '--endTime [endTime]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.startTime) {
          if (!validation.isValidISODateTime(args.options.startTime)) {
            return `'${args.options.startTime}' is not a valid ISO date string for option startTime.`;
          }
          if (new Date(args.options.startTime) > new Date()) {
            return 'startTime value cannot be in the future.';
          }
        }

        if (args.options.endTime) {
          if (!validation.isValidISODateTime(args.options.endTime)) {
            return `'${args.options.endTime}' is not a valid ISO date string for option endTime.`;
          }
          if (new Date(args.options.endTime) > new Date()) {
            return 'endTime value cannot be in the future.';
          }
        }

        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
          return 'endTime cannot be before startTime.';
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for option userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN for option userName.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('folderName', 'folderId', 'startTime', 'endTime', 'userId', 'userName');
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['folderId', 'folderName'],
        runsWhen: (args) => args.options.folderId || args.options.folderName
      },
      {
        options: ['userId', 'userName'],
        runsWhen: (args) => args.options.userId || args.options.userName
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'receivedDateTime'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.userId && !args.options.userName && accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken)) {
        throw 'You must specify either the userId or userName option when using app-only permissions.';
      }

      const userUrl = args.options.userId || args.options.userName ? `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}` : 'me';

      const folderId = await this.getFolderId(userUrl, args.options);
      const folderUrl: string = folderId ? `/mailFolders/${folderId}` : '';
      let requestUrl = `${this.resource}/v1.0/${userUrl}${folderUrl}/messages`;

      if (args.options.startTime || args.options.endTime) {
        const filters = [];

        if (args.options.startTime) {
          filters.push(`receivedDateTime ge ${args.options.startTime}`);
        }
        if (args.options.endTime) {
          filters.push(`receivedDateTime lt ${args.options.endTime}`);
        }

        if (filters.length > 0) {
          requestUrl += `?$filter=${filters.join(' and ')}`;
        }
      }

      const messages = await odata.getAllItems<Message>(requestUrl);
      await logger.log(messages);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getFolderId(userUrl: string, options: Options): Promise<string> {
    if (!options.folderId && !options.folderName) {
      return '';
    }

    if (options.folderId) {
      return options.folderId;
    }

    if (Outlook.wellKnownFolderNames.includes(options.folderName!.toLowerCase())) {
      return options.folderName!;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/${userUrl}/mailFolders?$filter=displayName eq '${formatting.encodeQueryParameter(options.folderName!)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: { id: string; }[] }>(requestOptions);

    if (response.value.length === 0) {
      throw `Folder with name '${options.folderName as string}' not found`;
    }

    if (response.value.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', response.value);
      const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple folders with name '${options.folderName!}' found.`, resultAsKeyValuePair);
      return result.id;
    }

    return response.value[0].id;
  }
}

export default new OutlookMessageListCommand();