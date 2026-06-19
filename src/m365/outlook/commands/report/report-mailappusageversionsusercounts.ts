import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

class OutlookReportMailAppUsageVersionsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILAPPUSAGEVERSIONSUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageVersionsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users by Outlook desktop version.';
  }
}

export default new OutlookReportMailAppUsageVersionsUserCountsCommand();