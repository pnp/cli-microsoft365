import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class YammerReportGroupsActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYGROUPCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityGroupCounts';
  }

  public get description(): string {
    return 'Gets the total number of groups that existed and how many included group conversation activity';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the total number of groups that existed and how many included group
    conversation activity for the last week
      ${commands.YAMMER_REPORT_GROUPSACTIVITYGROUPCOUNTS} --period D7

    Gets the total number of groups that existed and how many included group
    conversation activity for the last week and exports the report data in
    the specified path in text format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYGROUPCOUNTS} --period D7 --output text > "groupsactivitygroupcounts.txt"

    Gets the total number of groups that existed and how many included group
    conversation activity for the last week and exports the report data in
    the specified path in json format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYGROUPCOUNTS} --period D7 --output json > "groupsactivitygroupcounts.json"
`);
  }
}

module.exports = new YammerReportGroupsActivityGroupCountsCommand();