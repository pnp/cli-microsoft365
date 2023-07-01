import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageQuotaStatusMailboxCounts';
  }

  public get description(): string {
    return 'Gets the count of user mailboxes in each quota category';
  }
}

export default new OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand();