import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class YammerReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_DEVICEUSAGEUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the number of daily users by device type';
  }
}

module.exports = new YammerReportDeviceUsageUserCountsCommand();
