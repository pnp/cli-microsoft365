import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class YammerReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on the number of unique users who posted, read, and liked Yammer messages';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends on the number of unique users who posted, read, and liked
    Yammer messages for the last week
      ${commands.YAMMER_REPORT_ACTIVITYUSERCOUNTS} --period D7

    Gets the trends on the number of unique users who posted, read, and liked
    Yammer messages for the last week and exports the report data in the
    specified path in text format
      ${commands.YAMMER_REPORT_ACTIVITYUSERCOUNTS} --period D7 --output text > "activityusercounts.txt"

    Gets the trends on the number of unique users who posted, read, and liked
    Yammer messages for the last week and exports the report data in the
    specified path in json format
      ${commands.YAMMER_REPORT_ACTIVITYUSERCOUNTS} --period D7 --output json > "activityusercounts.json"
`);
  }
}

module.exports = new YammerReportActivityUserCountsCommand();
