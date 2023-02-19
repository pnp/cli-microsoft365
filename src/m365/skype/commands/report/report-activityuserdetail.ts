import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

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

export default new SkypeReportActivityUserDetailCommand();
