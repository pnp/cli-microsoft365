import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

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

module.exports = new OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand();