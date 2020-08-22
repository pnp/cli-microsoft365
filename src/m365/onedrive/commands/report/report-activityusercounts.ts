import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OneDriveReportActivityUserCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active OneDrive users';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trend in the number of active OneDrive users for the last week
      ${commands.REPORT_ACTIVITYUSERCOUNTS} --period D7

    Gets the trend in the number of active OneDrive users for the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_ACTIVITYUSERCOUNTS} --period D7 --output text > "activityusercounts.txt"

    Gets the trend in the number of active OneDrive users for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_ACTIVITYUSERCOUNTS} --period D7 --output json > "activityusercounts.json"
`);
  }
}

module.exports = new OneDriveReportActivityUserCountCommand();