import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { MeetingAttendanceReport } from '@microsoft/microsoft-graph-types';
import request, { CliRequestOptions } from '../../../../request.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  meetingId: string;
  id: string;
  userId?: string;
  userName?: string;
  email?: string;
}

class TeamsMeetingAttendancereportGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_ATTENDANCEREPORT_GET;
  }

  public get description(): string {
    return 'Gets attendance report for a given meeting';
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
        email: typeof args.options.email !== 'undefined'
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
        option: '--email [email]'
      },
      {
        option: '-m, --meetingId <meetingId>'
      },
      {
        option: '-i, --id <id>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID for option 'id'.`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for option 'userId'.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid UPN.`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName', 'email'],
      runsWhen: (args) => args.options.userId || args.options.userName || args.options.email
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
        await logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAccessToken ? `specific user ${args.options.userId || args.options.userName || args.options.email}` : 'currently logged in user'}.`);
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
      return await entraUser.getUserIdByUpn(options.userName);
    }
    return await entraUser.getUserIdByEmail(options.email!);
  }
}

export default new TeamsMeetingAttendancereportGetCommand();