import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

class OneDriveReportUsageAccountCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGEACCOUNTCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active OneDrive for Business sites';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }
}

export default new OneDriveReportUsageAccountCountsCommand();