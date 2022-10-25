import auth, { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from "../../../base/GraphCommand";
import commands from '../../commands';
import { odata } from '../../../../utils/odata';
import { Meeting } from '../Meeting';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userName?: string;
  isOrganizer: boolean;
}

class TeamsMeetingListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_LIST;
  }

  public get description(): string {
    return 'Retrieve all online meetings for a given organizer';
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'start', 'end'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userName: typeof args.options.userName !== 'undefined',
        isOrganizer: !!args.options.isOrganizer
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --userName [userName]'
      },
      {
        option: '--isOrganizer [isOrganizer]'
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
        if (!args.options.userName) {
          throw `The option 'userName' is required when retrieving meetings using app only credentials`;
        }
        requestUrl += `users/${args.options.userName}`;
      }
      else {
        if (this.verbose && args.options.userName) {
          throw `The option 'userName' cannot be set when retrieving meetings using delegated credentials`;
        }
        requestUrl += `me`;
      }

      requestUrl += '/events?$filter=isOrganizer eq true';
      /*if (args.options.isOrganizer) {
        requestUrl += '?$filter=isOrganizer eq true';
      }*/
      const res = await odata.getAllItems<Meeting>(requestUrl);
      const resFiltered = res.filter(y => y.isOnlineMeeting);
      if (!args.options.output || args.options.output === 'json') {
        logger.log(resFiltered);
      }
      else {
        //converted to text friendly output
        logger.log(resFiltered.map(i => {
          return {
            subject: i.subject,
            start: new Date(i.start.dateTime).toLocaleString(i.start.timeZone),
            end: new Date(i.end.dateTime).toLocaleString(i.end.timeZone)
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TeamsMeetingListCommand();