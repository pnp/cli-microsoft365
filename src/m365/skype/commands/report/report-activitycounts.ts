import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class SkypeReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.SKYPE_REPORT_ACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions';
  }
}

module.exports = new SkypeReportActivityCountsCommand();
