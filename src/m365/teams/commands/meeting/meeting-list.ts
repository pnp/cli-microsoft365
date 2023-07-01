import { Event } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  email?: string;
  startDateTime: string;
  endDateTime?: string;
  isOrganizer?: boolean;
}

class TeamsMeetingListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_LIST;
  }

  public get description(): string {
    return 'Retrieve all online meetings for a given user or shared mailbox';
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'start', 'end'];
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
        email: typeof args.options.email !== 'undefined',
        endDateTime: typeof args.options.endDateTime !== 'undefined',
        isOrganizer: !!args.options.isOrganizer
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
        option: '--startDateTime <startDateTime>'
      },
      {
        option: '--endDateTime [endDateTime]'
      },
      {
        option: '--isOrganizer'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidISODateTime(args.options.startDateTime)) {
          return `'${args.options.startDateTime}' is not a valid ISO date string`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid Guid`;
        }

        if (args.options.endDateTime && !validation.isValidISODateTime(args.options.endDateTime)) {
          return `'${args.options.startDateTime}' is not a valid ISO date string`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving meetings for ${isAppOnlyAccessToken ? 'specific user' : 'currently logged in user'}`);
      }

      let requestUrl = `${this.resource}/v1.0/`;
      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName && !args.options.email) {
          throw `The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`;
        }

        requestUrl += 'users/';
        if (args.options.userId) {
          requestUrl += args.options.userId;
        }
        else if (args.options.userName) {
          requestUrl += args.options.userName;
        }
        else if (args.options.email) {
          const userId = await aadUser.getUserIdByEmail(args.options.email);
          requestUrl += userId;
        }
      }
      else {
        if (args.options.userId || args.options.userName || args.options.email) {
          throw `The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`;
        }

        requestUrl += `me`;
      }

      requestUrl += `/events?$filter=start/dateTime ge '${args.options.startDateTime}'`;

      if (args.options.endDateTime) {
        requestUrl += ` and end/dateTime le '${args.options.endDateTime}'`;
      }

      if (args.options.isOrganizer) {
        requestUrl += ' and isOrganizer eq true';
      }

      const res = await odata.getAllItems<Event>(requestUrl);
      const resFiltered = res.filter(y => y.isOnlineMeeting);
      if (!args.options.output || !Cli.shouldTrimOutput(args.options.output)) {
        await logger.log(resFiltered);
      }
      else {
        //converted to text friendly output
        await logger.log(resFiltered.map(i => {
          return {
            subject: i.subject,
            start: i.start!.dateTime,
            end: i.end!.dateTime
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsMeetingListCommand();