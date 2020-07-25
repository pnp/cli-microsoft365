import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OutlookReportMailActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILACTIVITYCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityCounts';
  }

  public get description(): string {
    return 'Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization';
  }
}

module.exports = new OutlookReportMailActivityCountsCommand();