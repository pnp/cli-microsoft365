import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

class TeamsReportUserActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USERACTIVITYCOUNTS;
  }

  public get description(): string {
    return 'Gets the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityCounts';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }
}

export default new TeamsReportUserActivityCountsCommand();