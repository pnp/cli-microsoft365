import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

class OneDriveReportUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGESTORAGE;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageStorage';
  }

  public get description(): string {
    return 'Gets the trend on the amount of storage you are using in OneDrive for Business';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }
}

export default new OneDriveReportUsageStorageCommand();