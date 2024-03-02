import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { OnlineMeeting } from '@microsoft/microsoft-graph-types';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  startTime?: string;
  endTime?: string;
  subject?: string;
  participantUserNames?: string;
  organizerEmail?: string;
  recordAutomatically?: boolean;
}

class TeamsMeetingAddCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_ADD;
  }

  public get description(): string {
    return 'Creates a new online meeting';
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
        startTime: typeof args.options.startTime !== 'undefined',
        endTime: typeof args.options.endTime !== 'undefined',
        subject: typeof args.options.subject !== 'undefined',
        participantUserNames: typeof args.options.participantUserNames !== 'undefined',
        organizerEmail: typeof args.options.organizerEmail !== 'undefined',
        recordAutomatically: !!args.options.recordAutomatically
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --startTime [startTime]'
      },
      {
        option: '-e, --endTime [endTime]'
      },
      {
        option: '--subject [subject]'
      },
      {
        option: '-p, --participantUserNames [participantUserNames]'
      },
      {
        option: '--organizerEmail [organizerEmail]'
      },
      {
        option: '-r, --recordAutomatically'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.startTime && !validation.isValidISODateTime(args.options.startTime)) {
          return `'${args.options.startTime}' is not a valid ISO date string for startTime.`;
        }

        if (args.options.endTime && !validation.isValidISODateTime(args.options.endTime)) {
          return `'${args.options.endTime}' is not a valid ISO date string for endTime.`;
        }

        if (args.options.startTime && args.options.endTime && new Date(args.options.startTime) >= new Date(args.options.endTime)) {
          return 'The startTime value must be before endTime.';
        }

        if (args.options.startTime && new Date() >= new Date(args.options.startTime)) {
          return 'The startTime value must be in the future.';
        }

        if (args.options.endTime && new Date() >= new Date(args.options.endTime)) {
          return 'The endTime value must be in the future.';
        }

        if (args.options.participantUserNames) {
          const participants = args.options.participantUserNames.trim().toLowerCase().split(',').filter(e => e && e !== '');

          if (!participants || participants.length === 0 || participants.some(e => !validation.isValidUserPrincipalName(e))) {
            return `'${args.options.participantUserNames}' contains one or more invalid UPN.`;
          }
        }

        if (args.options.organizerEmail && !validation.isValidUserPrincipalName(args.options.organizerEmail)) {
          return `'${args.options.organizerEmail}' is not a valid email for organizerEmail.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)!;

      if (isAppOnlyAccessToken && !args.options.organizerEmail) {
        throw `The option 'organizerEmail' is required when creating a meeting using app only permissions`;
      }

      if (!isAppOnlyAccessToken && args.options.organizerEmail) {
        throw `The option 'organizerEmail' is not supported when creating a meeting using delegated permissions`;
      }

      const meeting = await this.createMeeting(logger, args.options);
      await logger.log(meeting);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Creates a new online meeting
   * @param logger 
   * @param options 
   * @returns MS Graph online meeting response
   */
  private async createMeeting(logger: Logger, options: Options): Promise<OnlineMeeting> {
    let requestUrl = `${this.resource}/v1.0/me`;

    if (options.organizerEmail) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving Organizer Id...`);
      }

      const organizerId = await entraUser.getUserIdByEmail(options.organizerEmail);
      requestUrl = `${this.resource}/v1.0/users/${organizerId}`;
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating the meeting...`);
    }

    const requestData: any = {};

    if (options.participantUserNames) {
      const attendees = options.participantUserNames.trim().toLowerCase().split(',').map(p => ({
        upn: p.trim()
      }));
      requestData.participants = { attendees };
    }

    if (options.startTime) {
      requestData.startDateTime = options.startTime;
    }

    if (options.endTime) {
      requestData.endDateTime = options.endTime;

      if (!options.startTime) {
        requestData.startDateTime = new Date().toISOString();
      }
    }

    if (options.subject) {
      requestData.subject = options.subject;
    }

    if (options.recordAutomatically !== undefined) {
      requestData.recordAutomatically = true;
    }

    const requestOptions: CliRequestOptions = {
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      url: `${requestUrl}/onlineMeetings`,
      data: requestData
    };

    return request.post<OnlineMeeting>(requestOptions);
  }
}

export default new TeamsMeetingAddCommand();