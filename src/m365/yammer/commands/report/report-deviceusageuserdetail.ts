import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class YammerReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer device usage by user';
  }
}

module.exports = new YammerReportDeviceUsageUserDetailCommand();