import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TeamsReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getTeamsDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams unique users by device type';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of Microsoft Teams unique users by device type for the last week
      ${commands.TEAMS_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS} --period D7

    Gets the number of Microsoft Teams unique users by device type for the last week
    and exports the report data in the specified path in text format
      ${commands.TEAMS_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS} --period D7 --output text > "deviceusagedistributionusercounts.txt"

    Gets the number of Microsoft Teams unique users by device type for the last week
    and exports the report data in the specified path in json format
      ${commands.TEAMS_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS} --period D7 --output json > "deviceusagedistributionusercounts.json"
`);
  }
}

module.exports = new TeamsReportDeviceUsageDistributionUserCountsCommand();