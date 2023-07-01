import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class OutlookReportMailAppUsageAppsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILAPPUSAGEAPPSUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageAppsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users per email app';
  }
}

export default new OutlookReportMailAppUsageAppsUserCountsCommand();