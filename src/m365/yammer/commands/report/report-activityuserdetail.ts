import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

class YammerReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer activity by user';
  }
}

module.exports = new YammerReportActivityUserDetailCommand();