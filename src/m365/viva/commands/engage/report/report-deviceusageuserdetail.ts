import DateAndPeriodBasedReport from '../../../../base/DateAndPeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEUSERDETAIL];
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer device usage by user';
  }
}

export default new YammerReportDeviceUsageUserDetailCommand();