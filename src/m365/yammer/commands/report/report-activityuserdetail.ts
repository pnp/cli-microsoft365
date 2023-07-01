import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class YammerReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer activity by user';
  }
}

export default new YammerReportActivityUserDetailCommand();