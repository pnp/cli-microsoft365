import { CalendarGroup } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { odata } from '../../../../utils/odata.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
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

class OutlookCalendarGroupListCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDARGROUP_LIST;
  }

  public get description(): string {
    return 'Retrieves calendar groups for a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both.'
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }

        const userIdentifier = args.options.userId ?? args.options.userName;

        if (this.verbose) {
          await logger.logToStderr(`Retrieving calendar groups for user '${userIdentifier}'...`);
        }

        const calendarGroups = await odata.getAllItems<CalendarGroup>(`${this.resource}/v1.0/users('${userIdentifier}')/calendarGroups`);
        await logger.log(calendarGroups);
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

          const userIdentifier = args.options.userId ?? args.options.userName;

          if (this.verbose) {
            await logger.logToStderr(`Retrieving calendar groups for user '${userIdentifier}'...`);
          }

          const calendarGroups = await odata.getAllItems<CalendarGroup>(`${this.resource}/v1.0/users('${userIdentifier}')/calendarGroups`);
          await logger.log(calendarGroups);
        }
        else {
          if (this.verbose) {
            await logger.logToStderr('Retrieving calendar groups for the signed-in user...');
          }

          const calendarGroups = await odata.getAllItems<CalendarGroup>(`${this.resource}/v1.0/me/calendarGroups`);
          await logger.log(calendarGroups);
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarGroupListCommand();
