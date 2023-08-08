import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class OutlookReportMailAppUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILAPPUSAGEUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users that connected to Exchange Online using any email app';
  }
}

export default new OutlookReportMailAppUsageUserCountsCommand();