import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class TeamsReportUserActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USERACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityUserDetail';
  }

  public get description(): string {
    return 'Get details about Microsoft Teams user activity by user.';
  }
}

export default new TeamsReportUserActivityUserDetailCommand();