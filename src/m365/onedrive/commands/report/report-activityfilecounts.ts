import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

class OneDriveReportActivityFileCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityFileCounts';
  }

  public get description(): string {
    return 'Gets the number of unique, licensed users that performed file interactions against any OneDrive account';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }
}

export default new OneDriveReportActivityFileCountCommand();