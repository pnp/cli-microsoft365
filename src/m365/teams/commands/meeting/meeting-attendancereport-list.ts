import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';
import { accessToken } from '../../../../utils/accessToken';

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
    const isAppOnlyAuth: boolean = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    if (isAppOnlyAuth && !args.options.userId && !args.options.userName && !args.options.email) {
      this.handleError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`);
    }
    else if (!isAppOnlyAuth && (args.options.userId || args.options.userName || args.options.email)) {
      this.handleError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`);
    }

    try {
      if (this.verbose) {
        logger.logToStderr(`Retrieving attendance report for ${isAppOnlyAuth ? 'specific user' : 'currently logged in user'}`);
      }

      let requestUrl = `${this.resource}/v1.0/`;
      if (isAppOnlyAuth) {
        requestUrl += 'users/';
        if (args.options.userId) {
          requestUrl += args.options.userId;
        }
        else if (args.options.userName) {
          requestUrl += args.options.userName;
        }
        else if (args.options.email) {
          const userId = await this.getUserId(args.options.email);
          requestUrl += userId;
        }
      }
      else {
        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings/${args.options.meetingId}/attendanceReports`;

      const res = await odata.getAllItems<any>(requestUrl);

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(email: string): Promise<string> {
    const options: AadUserGetCommandOptions = {
      email: email,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(AadUserGetCommand as Command, { options: { ...options, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.id;
  }
}

module.exports = new TeamsMeetingAttendancereportListCommand();