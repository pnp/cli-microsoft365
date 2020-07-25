import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageQuotaStatusMailboxCounts';
  }

  public get description(): string {
    return 'Gets the count of user mailboxes in each quota category';
  }
}

module.exports = new OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand();