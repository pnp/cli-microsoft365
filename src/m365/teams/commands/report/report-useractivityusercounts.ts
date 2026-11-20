import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  output: z.enum(['csv', 'json']).optional().alias('o'),
  period: z.enum(['D7', 'D30', 'D90', 'D180']).alias('p')
});

class TeamsReportUserActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USERACTIVITYUSERCOUNTS;
  }

  public get description(): string {
    return 'Gets the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityUserCounts';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }
}

export default new TeamsReportUserActivityUserCountsCommand();