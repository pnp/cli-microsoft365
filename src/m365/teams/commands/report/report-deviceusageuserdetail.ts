import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class TeamsReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public get description(): string {
    return 'Gets information about Microsoft Teams device usage by user';
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageUserDetail';
  }
}

export default new TeamsReportDeviceUsageUserDetailCommand();