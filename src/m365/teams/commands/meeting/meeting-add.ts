import auth from '../../../../Auth.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from "../../../base/GraphCommand.js";
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { OnlineMeeting } from '@microsoft/microsoft-graph-types';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  startTime: z.string()
    .refine(value => validation.isValidISODateTime(value), {
      message: 'The startTime value must be a valid ISO date string.'
    })
    .refine(value => new Date(value) > new Date(), {
      message: 'The startTime value must be in the future.'
    }).optional().alias('s'),
  endTime: z.string()
    .refine(value => validation.isValidISODateTime(value), {
      message: 'The endTime value must be a valid ISO date string.'
    })
    .refine(value => new Date(value) > new Date(), {
      message: 'The endTime value must be in the future.'
    }).optional().alias('e'),
  subject: z.string().optional(),
  participantUserNames: z.string()
    .superRefine((value, ctx) => {
      const invalidUserNames = validation.isValidUserPrincipalNameArray(value);
      if (invalidUserNames !== true) {
        ctx.addIssue({
          code: 'custom',
          message: `The following user principal names are invalid for the option 'participantUserNames': ${invalidUserNames}.`
        });
      }
    })
    .transform(value => {
      const userNames = value.trim().toLowerCase();
      // Command.action resolves standalone runtime tokens while they are strings.
      if (userNames === '@meusername') {
        return userNames;
      }

      return userNames.split(',').map(userName => userName.trim());
    })
    .optional().alias('p'),
  organizerEmail: z.string()
    .refine(value => validation.isValidUserPrincipalName(value), {
      message: 'The organizerEmail value must be a valid email.'
    }).optional(),
  recordAutomatically: z.boolean().optional().alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsMeetingAddCommand extends GraphCommand {
  public get name(): string {
    return commands.MEETING_ADD;
  }

  public get description(): string {
    return 'Creates a new online meeting';
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType {
    return schema.refine(options => !options.startTime || !options.endTime || new Date(options.startTime) < new Date(options.endTime), {
      message: 'The startTime value must be before endTime.',
      path: ['startTime']
    });
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
      const userNames = typeof options.participantUserNames === 'string'
        ? [options.participantUserNames.trim().toLowerCase()]
        : options.participantUserNames;
      const attendees = userNames.map(upn => ({
        upn
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