import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class OutlookReportMailAppUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_MAILAPPUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about which activities users performed on the various email apps';
  }
}

export default new OutlookReportMailAppUsageUserDetailCommand();