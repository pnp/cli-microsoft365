import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OutlookReportMailAppUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users that connected to Exchange Online using any email app';
  }
}

module.exports = new OutlookReportMailAppUsageUserCountsCommand();