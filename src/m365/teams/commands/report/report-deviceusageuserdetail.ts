import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

class TeamsReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.TEAMS_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public get description(): string {
    return 'Gets information about Microsoft Teams device usage by user';
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageUserDetail';
  }
}

module.exports = new TeamsReportDeviceUsageUserDetailCommand();