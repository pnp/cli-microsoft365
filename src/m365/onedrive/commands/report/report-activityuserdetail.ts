import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

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