import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

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

export default new OutlookReportMailboxUsageStorageCommand();