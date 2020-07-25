import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OutlookReportMailAppUsageVersionsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILAPPUSAGEVERSIONSUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageVersionsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users by Outlook desktop version.';
  }
}

module.exports = new OutlookReportMailAppUsageVersionsUserCountsCommand();