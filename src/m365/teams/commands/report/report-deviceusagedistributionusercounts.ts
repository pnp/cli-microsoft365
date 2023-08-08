import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class TeamsReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams unique users by device type';
  }
}

export default new TeamsReportDeviceUsageDistributionUserCountsCommand();