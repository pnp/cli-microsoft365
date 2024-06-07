import { Event } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

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
    return ['subject', 'startDateTime', 'endDateTime'];
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
          return `'${args.options.startDateTime}' is not a valid ISO date string for startDateTime.`;
        }

        if (args.options.endDateTime && !validation.isValidISODateTime(args.options.endDateTime)) {
          return `'${args.options.startDateTime}' is not a valid ISO date string for endDateTime.`;
        }

        if (args.options.startDateTime && args.options.endDateTime && args.options.startDateTime > args.options.endDateTime) {
          return 'startDateTime value must be before endDateTime.';
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID for userId.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `'${args.options.userName}' is not a valid UPN for userName.`;
        }

        if (args.options.email && !validation.isValidUserPrincipalName(args.options.email)) {
          return `'${args.options.email}' is not a valid UPN for email.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[this.resource].accessToken)!;
      if (isAppOnlyAccessToken && !args.options.userId && !args.options.userName && !args.options.email) {
        throw `The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`;
      }
      else if (!isAppOnlyAccessToken && (args.options.userId || args.options.userName || args.options.email)) {
        throw `The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`;
      }
      if (this.verbose) {
        await logger.logToStderr(`Retrieving meetings for user: ${args.options.userId || args.options.userName || args.options.email || accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[this.resource].accessToken)}...`);
      }

      const graphBaseUrl = await this.getGraphBaseUrl(args.options);
      const meetingUrls = await this.getMeetingJoinUrls(graphBaseUrl, args.options);
      const meetings = await this.getTeamsMeetings(logger, graphBaseUrl, meetingUrls);

      await logger.log(meetings);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  /**
   * Get the first part of the Graph API URL that contains the user information.
   */
  private async getGraphBaseUrl(options: Options): Promise<string> {
    let requestUrl = `${this.resource}/v1.0/`;

    if (options.userId || options.userName) {
      requestUrl += `users/${options.userId || options.userName}`;
    }
    else if (options.email) {
      const userId = await entraUser.getUserIdByEmail(options.email);
      requestUrl += `users/${userId}`;
    }
    else {
      requestUrl += 'me';
    }

    return requestUrl;
  }

  /**
   * Gets the meeting join urls for the specified user using calendar events.
   */
  private async getMeetingJoinUrls(graphBaseUrl: string, options: Options): Promise<string[]> {
    let requestUrl = graphBaseUrl;

    requestUrl += `/events?$filter=start/dateTime ge '${options.startDateTime}'`;
    if (options.endDateTime) {
      requestUrl += ` and end/dateTime lt '${options.endDateTime}'`;
    }
    if (options.isOrganizer) {
      requestUrl += ' and isOrganizer eq true';
    }
    requestUrl += '&$select=onlineMeeting';

    const items = await odata.getAllItems<Event>(requestUrl);
    const result = items.filter(i => i.onlineMeeting).map(i => i.onlineMeeting!.joinUrl!);

    return result;
  }

  private async getTeamsMeetings(logger: Logger, graphBaseUrl: string, meetingUrls: string[]): Promise<any[]> {
    const graphRelativeUrl = graphBaseUrl.replace(`${this.resource}/v1.0/`, '');
    let result: any[] = [];

    for (let i = 0; i < meetingUrls.length; i += 20) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving meetings ${i + 1}-${Math.min(i + 20, meetingUrls.length)}...`);
      }
      const batch = meetingUrls.slice(i, i + 20);
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/$batch`,
        headers: {
          accept: 'application/json',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          requests: batch.map((url, index) => ({
            id: i + index,
            method: 'GET',
            url: `${graphRelativeUrl}/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(url)}'`
          }))
        }
      };

      const requestResponse = await request.post<{ responses: { id: string; status: number; headers: any; body: any; }[] }>(requestOptions);

      for (const response of requestResponse.responses) {
        if (response.status === 200) {
          result.push(response.body.value[0]);
        }
        else {
          // Encountered errors where message was empty resulting in [object Object] error messages
          if (!response.body.error.message) {
            throw response.body.error.code;
          }
          throw response.body;
        }
      }
    }

    // Sort all meetings by start date
    result = result.sort((a, b) => a.startDateTime < b.startDateTime ? -1 : a.startDateTime > b.startDateTime ? 1 : 0);
    return result;
  }
}

export default new TeamsMeetingListCommand();