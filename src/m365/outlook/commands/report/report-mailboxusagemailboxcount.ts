import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class OutlookReportMailboxUsageMailboxCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILBOXUSAGEMAILBOXCOUNT;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageMailboxCounts';
  }

  public get description(): string {
    return 'Gets the total number of user mailboxes in your organization and how many are active each day of the reporting period.';
  }
}

export default new OutlookReportMailboxUsageMailboxCountCommand();
