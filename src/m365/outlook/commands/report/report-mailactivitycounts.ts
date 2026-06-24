import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

class OutlookReportMailActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityCounts';
  }

  public get description(): string {
    return 'Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization';
  }
}

export default new OutlookReportMailActivityCountsCommand();