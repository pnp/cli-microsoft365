import { CalendarGroup } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  name: z.string(),
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

class OutlookCalendarGroupAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDARGROUP_ADD;
  }

  public get description(): string {
    return 'Creates a calendar group for a user';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      let userUrl: string;
      let graphUserId: string;

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }

        graphUserId = (args.options.userId ?? args.options.userName)!;
        userUrl = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(graphUserId)}')`;

        if (this.verbose) {
          await logger.logToStderr(`Adding calendar group using application permissions for user '${graphUserId}'...`);
        }
      }
      else if (args.options.userId || args.options.userName) {
        const currentUserId = accessToken.getUserIdFromAccessToken(token);
        const currentUserName = accessToken.getUserNameFromAccessToken(token);
        const isOtherUser = (args.options.userId && args.options.userId !== currentUserId) ||
          (args.options.userName && args.options.userName.toLowerCase() !== currentUserName?.toLowerCase());

        if (isOtherUser) {
          const scopes = accessToken.getScopesFromAccessToken(token);
          const hasSharedScope = scopes.some(s => s === 'Calendars.ReadWrite.Shared');

          if (!hasSharedScope) {
            throw `To add calendar groups for other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.`;
          }
        }

        graphUserId = (args.options.userId ?? args.options.userName)!;
        userUrl = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(graphUserId)}')`;

        if (this.verbose) {
          await logger.logToStderr(`Adding calendar group using delegated permissions for user '${graphUserId}'...`);
        }
      }
      else {
        graphUserId = accessToken.getUserIdFromAccessToken(token);
        userUrl = `${this.resource}/v1.0/me`;

        if (this.verbose) {
          await logger.logToStderr('Adding calendar group for the signed-in user...');
        }
      }

      if (this.verbose) {
        await logger.logToStderr(`Creating calendar group '${args.options.name}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${userUrl}/calendarGroups`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          name: args.options.name
        }
      };

      const result = await request.post<CalendarGroup>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarGroupAddCommand();
