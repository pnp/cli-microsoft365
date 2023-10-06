import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { aadUser } from '../../../../utils/aadUser.js';
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
  participants?: string;
  organizerEmail?: string;
  recordAutomatically?: boolean;
}

class TeamsMeetingCreateCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_CREATE;
  }

  public get description(): string {
    return 'Create a new online meeting';
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'startDateTime', 'endDateTime', 'joinUrl', 'recordAutomatically'];
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
        participants: typeof args.options.participants !== 'undefined',
        organizerEmail: typeof args.options.organizerEmail !== 'undefined',
        recordAutomatically: !!args.options.recordAutomatically
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s --startTime [startTime]'
      },
      {
        option: '-e --endTime [endTime]'
      },
      {
        option: '-s --subject [subject]'
      },
      {
        option: '-p --participants [participants]'
      },
      {
        option: '--organizerEmail [organizerEmail]'
      },
      {
        option: '-r --recordAutomatically'
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
          return 'startTime value must be before endTime.';
        }
        if (args.options.endTime && !args.options.startTime) {
          return 'startTime should be specified when endTime is specified.';
        }
        if (args.options.participants) {
          if (args.options.participants.indexOf(',') === -1 && !validation.isValidUserPrincipalName(args.options.participants)) {
            return `${args.options.participants} contains invalid UPN.`;
          }
          const participants = args.options.participants.trim().toLowerCase().split(',').filter(e => e && e !== '');
          if (!participants || participants.length === 0 || participants.some(e => !validation.isValidUserPrincipalName(e))) {
            return `${args.options.participants} contains one or more invalid UPN.`;
          }
        }
        if (args.options.organizerEmail && !validation.isValidUserPrincipalName(args.options.organizerEmail)) {
          return `'${args.options.organizerEmail}' is not a valid email for organizerEmail.`;
        }
        return true;
      }
    );
  }

  /**
   * Executes the command
   * @param logger Logger instance
   * @param args Command arguments
   */
  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)!;
      if (isAppOnlyAccessToken && !args.options.organizerEmail) {
        throw `The option 'organizerEmail' is required when creating a meeting using app only permissions`;
      }
      const graphBaseUrl = await this.getGraphBaseUrl(args.options);
      const meeting = await this.createMeeting(logger, graphBaseUrl, args.options);
      await logger.log(meeting);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Gets the base MS Graph URL for the request
   * @param options 
   * @returns correct MS Graph URL for the request
   */
  private async getGraphBaseUrl(options: Options): Promise<string> {
    let requestUrl = `${this.resource}/v1.0/`;
    if (options.organizerEmail) {
      const organizerId = await aadUser.getUserIdByEmail(options.organizerEmail);
      requestUrl += `users/${organizerId}`;
    }
    else {
      requestUrl += 'me';
    }
    return requestUrl;
  }

  /**
   * Creates a new online meeting
   * @param logger 
   * @param graphBaseUrl 
   * @param options 
   * @returns MS Graph online meeting response
   */
  private async createMeeting(logger: Logger, graphBaseUrl: string, options: Options): Promise<any> {
    if (this.verbose) {
      logger.logToStderr(`Creation of a meeting...`);
    }
    const requestData: any = {};
    if (options.participants) {
      const attendees = options.participants.trim().toLowerCase().split(',').map(p => ({
        upn: p.trim()
      }));
      requestData.participants = { attendees };
    }
    if (options.startTime) {
      requestData.startDateTime = options.startTime;
    }
    if (options.endTime) {
      requestData.endDateTime = options.endTime;
    }
    if (options.subject) {
      requestData.subject = options.subject;
    }
    if (options.recordAutomatically !== undefined) {
      requestData.recordAutomatically = options.recordAutomatically;
    }
    const requestOption: CliRequestOptions = {
      headers: {
        accept: 'application/json',
        'content-type': 'application/json'
      },
      responseType: 'json',
      method: 'POST',
      url: `${graphBaseUrl}/onlineMeetings`,
      data: requestData
    };

    try {
      const requestResponse = await request.post<OnlineMeeting>(requestOption);
      return requestResponse;
    }
    catch (error: any) {
      if (error.response.status === 403) {
        throw `Forbidden. You do not have permission to perform this action. Please verify the command's details for more information.`;
      }

      throw error.message;
    }
  }
}

export default new TeamsMeetingCreateCommand();