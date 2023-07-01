import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class YammerReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer device usage by user';
  }
}

export default new YammerReportDeviceUsageUserDetailCommand();