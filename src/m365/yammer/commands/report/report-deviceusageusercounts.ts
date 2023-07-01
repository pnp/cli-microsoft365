import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

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

export default new YammerReportDeviceUsageUserCountsCommand();
