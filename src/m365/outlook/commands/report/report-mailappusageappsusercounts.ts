import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OutlookReportMailAppUsageAppsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILAPPUSAGEAPPSUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageAppsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users per email app';
  }
}

module.exports = new OutlookReportMailAppUsageAppsUserCountsCommand();