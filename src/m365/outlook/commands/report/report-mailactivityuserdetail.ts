import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class OutlookReportMailActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about email activity users have performed';
  }
}

module.exports = new OutlookReportMailActivityUserDetailCommand();