import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

class OneDriveReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about OneDrive activity by user';
  }
}

module.exports = new OneDriveReportActivityUserDetailCommand();