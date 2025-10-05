import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import GraphApplicationCommand from '../../../base/GraphApplicationCommand.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { odata } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional(),
  userName: z.string()
    .refine((val) => validation.isValidUserPrincipalName(val), {
      message: 'Invalid user principal name.'
    }).optional(),
  startDateTime: z.string()
    .refine((val) => {
      if (!validation.isValidISODateTime(val)) {
        return false;
      }
      const date = new Date(val);
      const maxDate = new Date();
      const minDate = new Date();
      minDate.setDate(maxDate.getDate() - 30);

      return date >= minDate && date <= maxDate;
    }, {
      message: 'Date must be a valid ISO date within the last 30 days and not in the future.'
    }).optional(),
  endDateTime: z.string()
    .refine((val) => {
      if (!validation.isValidISODateTime(val)) {
        return false;
      }
      const date = new Date(val);
      const maxDate = new Date();
      const minDate = new Date();
      minDate.setDate(maxDate.getDate() - 30);

      return date >= minDate && date <= maxDate;
    }, {
      message: 'Date must be a valid ISO date within the last 30 days and not in the future.'
    }).optional()
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class TeamsCallRecordListCommand extends GraphApplicationCommand {
  public get name(): string {
    return commands.CALLRECORD_LIST;
  }

  public get description(): string {
    return 'Lists all Teams calls within the tenant';
  }

  public defaultProperties(): string[] {
    return ['id', 'type', 'startDateTime', 'endDateTime'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine((options: Options) => [options.userId, options.userName].filter(o => o !== undefined).length <= 1, {
        error: 'Use one of the following options: userId or userName but not both.'
      })
      .refine((options: Options) => [options.startDateTime, options.endDateTime].filter(o => o !== undefined).length <= 1 || new Date(options.startDateTime!) < new Date(options.endDateTime!), {
        message: 'Value of startDateTime, must be before endDateTime.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving call records...`);
      }

      let apiUrl = `${this.resource}/v1.0/communications/callRecords`;
      const filters: string[] = [];
      if (args.options.userId || args.options.userName) {
        let userId = args.options.userId;

        if (args.options.userName) {
          userId = await entraUser.getUserIdByUpn(args.options.userName);
        }

        filters.push(`participants_v2/any(p:p/id eq '${userId}')`);
      }
      if (args.options.startDateTime) {
        filters.push(`startDateTime ge ${new Date(args.options.startDateTime).toISOString()}`);
      }
      if (args.options.endDateTime) {
        filters.push(`startDateTime lt ${new Date(args.options.endDateTime).toISOString()}`);
      }
      if (filters.length > 0) {
        apiUrl += `?$filter=${filters.join(' and ')}`;
      }

      const callRecords = await odata.getAllItems<any>(apiUrl);
      await logger.log(callRecords);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsCallRecordListCommand();