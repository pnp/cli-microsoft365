import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class TeamsReportUserActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USERACTIVITYCOUNTS;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityCounts';
  }
}

module.exports = new TeamsReportUserActivityCountsCommand();