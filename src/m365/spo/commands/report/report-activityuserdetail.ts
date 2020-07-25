import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class SpoReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_ACTIVITYUSERDETAIL}`;
  }

  public get description(): string {
    return 'Gets details about SharePoint activity by user';
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityUserDetail';
  }
}

module.exports = new SpoReportActivityUserDetailCommand();