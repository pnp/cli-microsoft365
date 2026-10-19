import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command, { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import entraUserGetCommand, { Options as EntraUserGetCommandOptions } from '../../../entra/commands/user/user-get.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The userId value must be a valid GUID.'
    }).optional().alias('u'),
  userName: z.string().optional().alias('n'),
  email: z.string().optional(),
  meetingId: z.string().alias('m')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingAttendancereportListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_ATTENDANCEREPORT_LIST;
  }

  public get description(): string {
    return 'Lists all attendance reports for a given meeting';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'totalParticipantCount'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
    if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
      this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`);
    }
    else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
      this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`);
    }

    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
      }

      let requestUrl = `${this.resource}/v1.0/`;
      if (isAppOnlyAccessToken) {
        requestUrl += 'users/';
        if (args.options.userId) {
          requestUrl += args.options.userId;
        }
        else {
          const userId = await this.getUserId(args.options.userName, args.options.email);
          requestUrl += userId;
        }
      }
      else {
        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings/${args.options.meetingId}/attendanceReports`;

      const res = await odata.getAllItems<any>(requestUrl);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(userName?: string, email?: string): Promise<string> {
    const options: EntraUserGetCommandOptions = {
      email: email,
      userName: userName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(entraUserGetCommand as Command, { options: { ...options, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.id;
  }
}

export default new TeamsMeetingAttendancereportListCommand();