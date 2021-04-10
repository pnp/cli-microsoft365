import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

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

module.exports = new TeamsReportUserActivityUserDetailCommand();