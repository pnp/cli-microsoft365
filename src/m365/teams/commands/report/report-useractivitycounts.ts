import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class TeamsReportUserActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityCounts';
  }
}

module.exports = new TeamsReportUserActivityCountsCommand();