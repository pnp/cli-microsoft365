import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).optional().alias('p'),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/, 'The supported date format is YYYY-MM-DD').optional().alias('d')
});
declare type Options = z.infer<typeof options>;

class OneDriveReportUsageAccountDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGEACCOUNTDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountDetail';
  }

  public get description(): string {
    return 'Gets details about OneDrive usage by account';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine((opts: Options) => opts.period !== undefined || opts.date !== undefined, {
        message: 'Specify period or date, one is required.',
        params: { customCode: 'optionSet', options: ['period', 'date'] }
      })
      .refine((opts: Options) => !(opts.period !== undefined && opts.date !== undefined), {
        message: 'Specify period or date but not both.'
      });
  }
}

export default new OneDriveReportUsageAccountDetailCommand();