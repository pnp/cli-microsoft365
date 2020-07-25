import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class SkypeReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.SKYPE_REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Skype for Business activity by user';
  }
}

module.exports = new SkypeReportActivityUserDetailCommand();
