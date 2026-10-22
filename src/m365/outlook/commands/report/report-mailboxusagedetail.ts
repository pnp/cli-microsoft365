import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

class OutlookReportMailboxUsageDetailCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILBOXUSAGEDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageDetail';
  }

  public get description(): string {
    return 'Gets details about mailbox usage';
  }
}

export default new OutlookReportMailboxUsageDetailCommand();