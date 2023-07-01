import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import aadUserGetCommand, { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  email?: string;
  meetingId: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid Guid`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
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
    const options: AadUserGetCommandOptions = {
      email: email,
      userName: userName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(aadUserGetCommand as Command, { options: { ...options, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.id;
  }
}

export default new TeamsMeetingAttendancereportListCommand();