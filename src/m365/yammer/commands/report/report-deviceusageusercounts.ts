import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class YammerReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the number of daily users by device type';
  }
}

module.exports = new YammerReportDeviceUsageUserCountsCommand();
