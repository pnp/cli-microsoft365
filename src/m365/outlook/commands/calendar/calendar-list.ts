import { Calendar } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { odata } from '../../../../utils/odata.js';

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
  calendarGroupId: z.string().optional(),
  calendarGroupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarListCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDAR_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of all calendars of a user or a group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName].filter(x => x !== undefined).length === 1, {
        error: 'Specify either userId or userName, but not both'
      })
      .refine(options => !(options.calendarGroupId && options.calendarGroupName), {
        error: 'Specify either calendarGroupId or calendarGroupName, but not both'
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving calendars for ${args.options.userId ?? args.options.userName}...`);
    }

    try {
      const userIdentifier = args.options.userId ?? args.options.userName;
      let calendarGroupId = args.options.calendarGroupId;

      if (args.options.calendarGroupName) {
        const group = await calendarGroup.getUserCalendarGroupByName(userIdentifier!, args.options.calendarGroupName, 'id');
        calendarGroupId = group.id;
      }

      const url = `${this.resource}/v1.0/users('${userIdentifier}')/${calendarGroupId ? `calendarGroups/${calendarGroupId}/` : ''}calendars`;
      const calendars = await odata.getAllItems<Calendar>(url);

      await logger.log(calendars);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarListCommand();