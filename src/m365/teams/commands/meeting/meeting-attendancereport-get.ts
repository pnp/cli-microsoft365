import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { MeetingAttendanceReport } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The userId value must be a valid GUID.'
    }).optional().alias('u'),
  userName: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The userName value must be a valid user principal name (UPN).'
    }).optional().alias('n'),
  email: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The email value must be a valid email.'
    }).optional(),
  meetingId: z.string().alias('m'),
  id: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The id value must be a valid GUID.'
    }).alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingAttendancereportGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_ATTENDANCEREPORT_GET;
  }

  public get description(): string {
    return 'Gets attendance report for a given meeting';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType {
    return schema.refine(options => [options.userId, options.userName, options.email].filter(value => value !== undefined).length <= 1, {
      message: 'Specify either userId, userName or email, but not multiple.',
      params: {
        customCode: 'optionSet',
        options: ['userId', 'userName', 'email']
      }
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
        throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions.`;
      }
      else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
        throw `The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance report using delegated permissions.`;
      }

      if (this.verbose) {
        await logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAccessToken ? `specific user ${args.options.userId || args.options.userName || args.options.email}.` : 'currently logged in user'}.`);
      }

      let userUrl = '';
      if (isAppOnlyAccessToken) {
        const userId = await this.getUserId(args.options);
        userUrl += `users/${userId}`;
      }
      else {
        userUrl += 'me';
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/${userUrl}/onlineMeetings/${args.options.meetingId}/attendanceReports/${args.options.id}?$expand=attendanceRecords`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const attendanceReport = await request.get<MeetingAttendanceReport>(requestOptions);
      await logger.log(attendanceReport);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(options: Options): Promise<string> {
    if (options.userId) {
      return options.userId;
    }

    if (options.userName) {
      return entraUser.getUserIdByUpn(options.userName);
    }

    return entraUser.getUserIdByEmail(options.email!);
  }
}

export default new TeamsMeetingAttendancereportGetCommand();