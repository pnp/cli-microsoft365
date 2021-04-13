import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

class SkypeReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Skype for Business activity by user';
  }
}

module.exports = new SkypeReportActivityUserDetailCommand();
