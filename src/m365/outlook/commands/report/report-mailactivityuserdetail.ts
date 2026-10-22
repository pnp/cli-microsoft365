import DateAndPeriodBasedReport, { dateAndPeriodBasedReportOptions } from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

export const options = dateAndPeriodBasedReportOptions;

class OutlookReportMailActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about email activity users have performed';
  }
}

export default new OutlookReportMailActivityUserDetailCommand();