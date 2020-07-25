import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class TeamsReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams unique users by device type';
  }
}

module.exports = new TeamsReportDeviceUsageDistributionUserCountsCommand();