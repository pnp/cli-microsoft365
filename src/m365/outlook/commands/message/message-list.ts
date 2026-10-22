import { Message } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
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
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  folderId: z.string().optional(),
  folderName: z.string().optional(),
  startTime: z.string()
    .refine(startTime => validation.isValidISODateTime(startTime), {
      error: e => `'${e.input}' is not a valid ISO date string for option startTime.`
    })
    .refine(startTime => new Date(startTime) <= new Date(), {
      error: 'startTime value cannot be in the future.'
    })
    .optional(),
  endTime: z.string()
    .refine(endTime => validation.isValidISODateTime(endTime), {
      error: e => `'${e.input}' is not a valid ISO date string for option endTime.`
    })
    .refine(endTime => new Date(endTime) <= new Date(), {
      error: 'endTime value cannot be in the future.'
    })
    .optional(),
  userId: z.string()
    .refine(userId => validation.isValidGuid(userId), {
      error: e => `${e.input} is not a valid GUID for option userId.`
    }).optional(),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `${e.input} is not a valid UPN for option userName.`
    }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_LIST;
  }

  public get description(): string {
    return 'Gets all mail messages from the specified folder';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.folderId && options.folderName), {
        error: 'Specify either folderId or folderName, but not both',
        params: {
          customCode: 'optionSet',
          options: ['folderId', 'folderName']
        }
      })
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      })
      .refine(options => !(options.startTime && options.endTime && new Date(options.startTime) >= new Date(options.endTime)), {
        error: 'startTime must be before endTime.'
      });
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
      let requestUrl = `${this.resource}/v1.0/${userUrl}${folderUrl}/messages?$top=100`;

      if (args.options.startTime || args.options.endTime) {
        const filters = [];

        if (args.options.startTime) {
          filters.push(`receivedDateTime ge ${args.options.startTime}`);
        }
        if (args.options.endTime) {
          filters.push(`receivedDateTime lt ${args.options.endTime}`);
        }

        if (filters.length > 0) {
          requestUrl += `&$filter=${filters.join(' and ')}`;
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