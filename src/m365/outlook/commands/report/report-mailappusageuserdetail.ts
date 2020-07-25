import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class OutlookReportMailAppUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about which activities users performed on the various email apps';
  }
}

module.exports = new OutlookReportMailAppUsageUserDetailCommand();