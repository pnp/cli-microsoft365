import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class TeamsReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_DEVICEUSAGEUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams daily unique users by device type';
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageUserCounts';
  }
}

module.exports = new TeamsReportDeviceUsageUserCountsCommand();