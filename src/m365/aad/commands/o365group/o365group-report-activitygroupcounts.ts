import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class O365GroupReportActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.O365GROUP_REPORT_ACTIVITYGROUPCOUNTS}`;
  }

  public get description(): string {
    return 'Get the daily total number of groups and how many of them were active based on email conversations, Yammer posts, and SharePoint file activities';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityGroupCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get the daily total number of groups and how many of them were active based
    on activities for the last week
      ${commands.O365GROUP_REPORT_ACTIVITYGROUPCOUNTS} --period D7

    Get the daily total number of groups and how many of them were active based
    on activities for the last week and exports the report data in the specified
    path in text format
      ${commands.O365GROUP_REPORT_ACTIVITYGROUPCOUNTS} --period D7 --output text > "o365groupactivitygroupcounts.txt"

    Get the daily total number of groups and how many of them were active based
    on activities for the last week and exports the report data in the specified
    path in json format
      ${commands.O365GROUP_REPORT_ACTIVITYGROUPCOUNTS} --period D7 --output json > "o365groupactivitygroupcounts.json"
`);
  }
}

module.exports = new O365GroupReportActivityGroupCountsCommand();