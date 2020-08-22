import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class YammerReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Yammer messages posted, read, and liked in groups';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of Yammer messages posted, read, and liked in groups for
    the last week
      ${commands.YAMMER_REPORT_GROUPSACTIVITYCOUNTS} --period D7

    Gets the number of Yammer messages posted, read, and liked in groups for
    the last week and exports the report data in the specified path in text
    format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYCOUNTS} --period D7 --output text > "groupsactivitycounts.txt"

    Gets the number of Yammer messages posted, read, and liked in groups for
    the last week and exports the report data in the specified path in json
    format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYCOUNTS} --period D7 --output json > "groupsactivitycounts.json"
`);
  }
}

module.exports = new YammerReportGroupsActivityCountsCommand();
