import auth, { Auth } from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import Command from '../../../../Command';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { Meeting } from '../Meeting';
import { validation } from '../../../../utils/validation';
import * as AadUserGetCommand from '../../../aad/commands/user/user-get';
import { Options as AadUserGetCommandOptions } from '../../../aad/commands/user/user-get';

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
      const isAppOnlyAuth: boolean | undefined = Auth.isAppOnlyAuth(auth.service.accessTokens[this.resource].accessToken);
      if (this.verbose) {
        logger.logToStderr(`Retrieving meetings for ${isAppOnlyAuth ? 'specific user' : 'currently logged in user'}`);
      }

      let requestUrl = `${this.resource}/v1.0/`;
      if (isAppOnlyAuth) {
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
          const userId = await this.getUserId(args.options.email);
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

      const res = await odata.getAllItems<Meeting>(requestUrl);
      const resFiltered = res.filter(y => y.isOnlineMeeting);
      if (!args.options.output || !Cli.shouldTrimOutput(args.options.output)) {
        logger.log(resFiltered);
      }
      else {
        //converted to text friendly output
        logger.log(resFiltered.map(i => {
          return {
            subject: i.subject,
            start: i.start.dateTime,
            end: i.end.dateTime
          };
        }));
      }
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

module.exports = new TeamsMeetingListCommand();