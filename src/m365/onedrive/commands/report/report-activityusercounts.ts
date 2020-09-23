import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class OneDriveReportActivityUserCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active OneDrive users';
  }
}

module.exports = new OneDriveReportActivityUserCountCommand();