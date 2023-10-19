import PeriodBasedReport from '../../../../base/PeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEUSERCOUNTS];
  }
  
  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the number of daily users by device type';
  }
}

export default new YammerReportDeviceUsageUserCountsCommand();
