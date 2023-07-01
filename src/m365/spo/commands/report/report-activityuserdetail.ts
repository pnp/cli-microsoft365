import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class SpoReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERDETAIL;
  }

  public get description(): string {
    return 'Gets details about SharePoint activity by user';
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityUserDetail';
  }
}

export default new SpoReportActivityUserDetailCommand();