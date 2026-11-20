import { Event } from '@microsoft/microsoft-graph-types';
import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(value => validation.isValidGuid(value), {
      message: 'The userId value must be a valid GUID.'
    }).optional().alias('u'),
  userName: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The userName value must be a valid user principal name (UPN).'
    }).optional().alias('n'),
  email: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The email value must be a valid email.'
    }).optional(),
  startDateTime: z.string()
    .refine(value => validation.isValidISODateTime(value), {
      error: issue => `'${issue.input}' is not a valid ISO date string for startDateTime.`
    }),
  endDateTime: z.string()
    .refine(value => validation.isValidISODateTime(value), {
      error: issue => `'${issue.input}' is not a valid ISO date string for endDateTime.`
    }).optional(),
  isOrganizer: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingListCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_LIST;
  }

  public get description(): string {
    return 'Retrieves all online meetings for a given user or shared mailbox';
  }

  public defaultProperties(): string[] | undefined {
    return ['subject', 'startDateTime', 'endDateTime'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType {
    return schema.refine(options => !options.endDateTime || options.startDateTime <= options.endDateTime, {
      message: 'startDateTime value must be before endDateTime.',
      path: ['startDateTime']
    });
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