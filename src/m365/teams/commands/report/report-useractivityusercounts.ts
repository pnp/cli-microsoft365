import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class TeamsReportUserActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USERACTIVITYUSERCOUNTS;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityUserCounts';
  }
}

export default new TeamsReportUserActivityUserCountsCommand();