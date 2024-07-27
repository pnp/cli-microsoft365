import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Viva Engage activity by user';
  }
}

export default new VivaEngageReportActivityUserDetailCommand();