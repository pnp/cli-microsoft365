import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of daily users by device type for
    the last week
      ${commands.YAMMER_REPORT_DEVICEUSAGEUSERCOUNTS} --period D7

    Gets the number of daily users by device type for the last week and exports
    the report data in the specified path in text format
      ${commands.YAMMER_REPORT_DEVICEUSAGEUSERCOUNTS} --period D7 --output text > "deviceusageusercounts.txt"

    Gets the number of daily users by device type for the last week and exports
    the report data in the specified path in json format
      ${commands.YAMMER_REPORT_DEVICEUSAGEUSERCOUNTS} --period D7 --output json > "deviceusageusercounts.json"
`);
  }
}

module.exports = new YammerReportDeviceUsageUserCountsCommand();
