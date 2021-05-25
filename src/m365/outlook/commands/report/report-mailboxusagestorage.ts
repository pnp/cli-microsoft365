import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class OutlookReportMailboxUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILBOXUSAGESTORAGE;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageStorage';
  }

  public get description(): string {
    return 'Gets the amount of mailbox storage used in your organization';
  }
}

module.exports = new OutlookReportMailboxUsageStorageCommand();