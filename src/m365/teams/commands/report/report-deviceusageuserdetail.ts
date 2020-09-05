import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TeamsReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_DEVICEUSAGEUSERDETAIL}`;
  }

  public get description(): string {
    return 'Gets information about Microsoft Teams device usage by user';
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageUserDetail';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    As this report is only available for the past 28 days, date parameter value
    should be a date from that range.

  Examples:

    Gets information about Microsoft Teams device usage by user for the last week
      m365 ${this.name} --period D7

    Gets information about Microsoft Teams device usage by user for
    May 1, 2019
      m365 ${this.name} --date 2019-05-01

    Gets information about Microsoft Teams device usage by user for the last week 
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "deviceusageuserdetails.txt"

    Gets information about Microsoft Teams device usage by user for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "deviceusageuserdetails.json"
`);
  }
}

module.exports = new TeamsReportDeviceUsageUserDetailCommand();