import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { validation } from '../../../../utils/validation.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';
import { calendar } from '../../../../utils/calendar.js';
import { cli } from '../../../../cli/cli.js';
import request, { CliRequestOptions } from '../../../../request.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i').optional(),
  name: z.string().alias('n').optional(),
  userId: z.string()
    .refine(userId => validation.isValidGuid(userId), {
      error: e => `'${e.input}' is not a valid GUID.`
    }).optional(),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `'${e.input}' is not a valid UPN.`
    }).optional(),
  calendarGroupId: z.string().optional(),
  calendarGroupName: z.string().optional(),
  permanent: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDAR_REMOVE;
  }

  public get description(): string {
    return 'Removes the calendar of a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.id, options.name].filter(x => x !== undefined).length === 1, {
        error: 'Specify either id or name, but not both'
      })
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both'
      })
      .refine(options => [options.calendarGroupId, options.calendarGroupName].filter(x => x !== undefined).length !== 2, {
        error: 'Do not specify both calendarGroupId and calendarGroupName'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeCalendar = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr('Getting calendar...');
      }

      try {
        const userIdentifier = args.options.userId ?? args.options.userName;
        let calendarGroupId = args.options.calendarGroupId;

        if (args.options.calendarGroupName) {
          const group = await calendarGroup.getUserCalendarGroupByName(userIdentifier!, args.options.calendarGroupName, 'id');
          calendarGroupId = group.id;
        }

        let calendarId = args.options.id;
        if (args.options.name) {
          const result = await calendar.getUserCalendarByName(userIdentifier!, args.options.name!, calendarGroupId, 'id');
          calendarId = result.id;
        }

        let url = `${this.resource}/v1.0/users('${userIdentifier}')/${calendarGroupId ? `calendarGroups/${calendarGroupId}/` : ''}calendars/${calendarId}`;
        if (args.options.permanent) {
          url += '/permanentDelete';
        }
        const requestOptions: CliRequestOptions = {
          url: url,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        if (args.options.permanent) {
          await request.post(requestOptions);
        }
        else {
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeCalendar();
    }
    else {
      const calendarIdentifier = args.options.id ?? args.options.name;
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove calendar '${calendarIdentifier}'?` });
      if (result) {
        await removeCalendar();
      }
    }
  }
}

export default new OutlookCalendarRemoveCommand();
