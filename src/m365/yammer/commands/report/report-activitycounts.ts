import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class YammerReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_ACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends on the amount of Yammer activity in your organization by how
    many messages were posted, read, and liked for the last week
      ${commands.YAMMER_REPORT_ACTIVITYCOUNTS} --period D7

    Gets the trends on the amount of Yammer activity in your organization by how
    many messages were posted, read, and liked for the last week and exports the
    report data in the specified path in text format
      ${commands.YAMMER_REPORT_ACTIVITYCOUNTS} --period D7 --output text > "activitycounts.txt"

    Gets the trends on the amount of Yammer activity in your organization by how
    many messages were posted, read, and liked for the last week and exports the
    report data in the specified path in json format
      ${commands.YAMMER_REPORT_ACTIVITYCOUNTS} --period D7 --output json > "activitycounts.json"
`);
  }
}

module.exports = new YammerReportActivityCountsCommand();
