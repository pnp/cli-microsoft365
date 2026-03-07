import { Calendar } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.string()
    .refine(userId => validation.isValidGuid(userId), {
      error: e => `'${e.input}' is not a valid GUID.`
    }).optional(),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `'${e.input}' is not a valid UPN.`
    }).optional(),
  name: z.string(),
  calendarGroupId: z.string().optional(),
  calendarGroupName: z.string().optional(),
  color: z.enum(['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'maxColor']).optional().default('auto'),
  defaultOnlineMeetingProvider: z.enum(['none', 'teamsForBusiness']).optional().default('teamsForBusiness'),
  default: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDAR_ADD;
  }

  public get description(): string {
    return 'Creates a new calendar for a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {

      const userIdentifier = args.options.userId ?? args.options.userName;

      let requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/`;

      if (args.options.calendarGroupId || args.options.calendarGroupName) {
        let calendarGroupId = args.options.calendarGroupId;
        if (args.options.calendarGroupName) {
          const group = await calendarGroup.getUserCalendarGroupByName(userIdentifier!, args.options.calendarGroupName, 'id');
          calendarGroupId = group.id;
        }
        requestUrl += `calendarGroups/${calendarGroupId}/calendars`;
      }
      else {
        requestUrl += 'calendars';
      }

      if (args.options.verbose) {
        await logger.logToStderr(`Creating a calendar for the user ${userIdentifier}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          name: args.options.name,
          color: args.options.color,
          defaultOnlineMeetingProvider: args.options.defaultOnlineMeetingProvider,
          isDefaultCalendar: args.options.default
        }
      };

      const result = await request.post<Calendar>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarAddCommand();
