import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

class OutlookReportMailActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityUserCounts';
  }

  public get description(): string {
    return 'Enables you to understand trends on the number of unique users who are performing email activities like send, read, and receive';
  }
}

export default new OutlookReportMailActivityUserCountsCommand();