import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import { accessToken } from '../../../../utils/accessToken';
import { aadUser } from '../../../../utils/aadUser';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  email?: string;
  meetingId: string;
}

class TeamsMeetingTranscriptListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_TRANSCRIPT_LIST;
  }

  public get description(): string {
    return 'Lists all transcripts for a given meeting';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'createdDateTime'];
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
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid Guid`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `${args.options.email} is not a valid email`;
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
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        logger.logToStderr(`Retrieving transcript list for the given meeting...`);
      }

      let requestUrl: string = `${this.resource}/beta/`;
      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName && !args.options.email) {
          throw `The option 'userId', 'userName' or 'email' is required when retrieving meeting transcripts list using app only permissions`;
        }

        requestUrl += 'users/';
        if (args.options.userId) {
          requestUrl += args.options.userId;
        }
        else if (args.options.userName) {
          requestUrl += args.options.userName;
        }
        else if (args.options.email) {
          const userId: string = await aadUser.getUserIdByEmail(args.options.email!);
          requestUrl += userId;
        }
      }
      else {
        if (args.options.userId || args.options.userName || args.options.email) {
          throw `The options 'userId', 'userName' and 'email' cannot be used while retrieving meeting transcripts using delegated permissions`;
        }

        requestUrl += `me`;
      }

      requestUrl += `/onlineMeetings/${args.options.meetingId}/transcripts`;
      const res = await odata.getAllItems<any>(requestUrl);

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMeetingTranscriptListCommand();