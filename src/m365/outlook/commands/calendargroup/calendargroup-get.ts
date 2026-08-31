import { CalendarGroup } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  name: z.string().optional(),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarGroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDARGROUP_GET;
  }

  public get description(): string {
    return 'Retrieves a calendar group for a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => !(o.id && o.name), {
        error: 'Specify either id or name, but not both.'
      })
      .refine(o => Boolean(o.id || o.name), {
        error: 'Specify either id or name.'
      })
      .refine(o => !(o.userId && o.userName), {
        error: 'Specify either userId or userName, but not both.'
      });
  }


  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      let userIdentifier: string | undefined = undefined;
      if (args.options.userId || args.options.userName) {
        userIdentifier = args.options.userId ?? args.options.userName;
      }

      const encodedUserIdentifier: string | undefined = userIdentifier
        ? formatting.encodeQueryParameter(userIdentifier)
        : undefined;

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }
      }
      else {
        if (args.options.userId || args.options.userName) {
          const currentUserId = accessToken.getUserIdFromAccessToken(token);
          const currentUserName = accessToken.getUserNameFromAccessToken(token);

          const isOtherUser = (args.options.userId && args.options.userId !== currentUserId) ||
            (args.options.userName && args.options.userName.toLowerCase() !== currentUserName?.toLowerCase());

          if (isOtherUser) {
            const scopes = accessToken.getScopesFromAccessToken(token);
            const hasSharedScope = scopes.some(s => s === 'Calendars.Read.Shared' || s === 'Calendars.ReadWrite.Shared');
            if (!hasSharedScope) {
              throw `To retrieve calendar groups of other users, the Entra ID application used for authentication must have either the Calendars.Read.Shared or Calendars.ReadWrite.Shared delegated permission assigned.`;
            }
          }
        }
      }

      const getCalendarGroupByName = async (calendarGroupName: string): Promise<CalendarGroup> => {
        const userPath = encodedUserIdentifier ? `users('${encodedUserIdentifier}')` : 'me';
        const calendarGroups = await odata.getAllItems<CalendarGroup>(
          `${this.resource}/v1.0/${userPath}/calendarGroups?$filter=name eq '${formatting.encodeQueryParameter(calendarGroupName)}'`
        );

        if (calendarGroups.length === 0) {
          throw `The specified calendar group '${calendarGroupName}' does not exist.`;
        }

        return calendarGroups[0];
      };

      if (args.options.name) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving calendar group '${args.options.name}'...`);
        }
        const result = await getCalendarGroupByName(args.options.name);
        await logger.log(result);
      }
      else {
        const calendarGroupId = args.options.id!;
        const userPath = encodedUserIdentifier ? `users('${encodedUserIdentifier}')` : 'me';
        const requestUrl = `${this.resource}/v1.0/${userPath}/calendarGroups/${calendarGroupId}`;

        if (this.verbose) {
          await logger.logToStderr(`Retrieving calendar group '${calendarGroupId}'...`);
        }

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const result = await request.get<CalendarGroup>(requestOptions);
        await logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarGroupGetCommand();

