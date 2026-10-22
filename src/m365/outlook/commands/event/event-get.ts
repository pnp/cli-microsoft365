import { Event } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { calendar } from '../../../../utils/calendar.js';
import { formatting } from '../../../../utils/formatting.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  calendarId: z.string().optional(),
  calendarName: z.string().optional(),
  timeZone: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookEventGetCommand extends GraphCommand {
  public get name(): string {
    return commands.EVENT_GET;
  }

  public get description(): string {
    return `Retrieve an event from a specific calendar of a user`;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName].filter(x => x !== undefined).length === 1, {
        error: 'Specify either userId or userName, but not both'
      })
      .refine(options => !(options.calendarId && options.calendarName), {
        error: 'Specify either calendarId or calendarName, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const userIdentifier = args.options.userId ?? args.options.userName;
    if (this.verbose) {
      await logger.logToStderr(`Retrieving event ${args.options.id} for user ${userIdentifier}...`);
    }

    let calendarId = args.options.calendarId;
    if (args.options.calendarName) {
      calendarId = (await calendar.getUserCalendarByName(userIdentifier!, args.options.calendarName))!.id;
    }

    let requestUrl: string = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(userIdentifier!)}')`;

    if (calendarId) {
      requestUrl += `/calendars/${calendarId}/events/${args.options.id}`;
    }
    else {
      requestUrl += `/events/${args.options.id}`;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    if (args.options.timeZone) {
      requestOptions.headers!.Prefer = `outlook.timezone="${args.options.timeZone}"`;
    }

    try {
      const result = await request.get<Event>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookEventGetCommand();