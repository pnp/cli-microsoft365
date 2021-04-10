import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class YammerReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Gets the number of users by device type';
  }
}

module.exports = new YammerReportDeviceUsageDistributionUserCountsCommand();

