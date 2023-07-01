import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class TeamsReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEUSERCOUNTS;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams daily unique users by device type';
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageUserCounts';
  }
}

export default new TeamsReportDeviceUsageUserCountsCommand();