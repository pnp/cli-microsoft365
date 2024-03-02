import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
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
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving transcript list for the given meeting...`);
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
          if (this.verbose) {
            await logger.logToStderr(`Getting user ID for user with email '${args.options.email}'.`);
          }
          const userId: string = await entraUser.getUserIdByEmail(args.options.email!);
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

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMeetingTranscriptListCommand();