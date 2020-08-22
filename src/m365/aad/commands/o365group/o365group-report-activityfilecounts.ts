import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class O365GroupReportActivityFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYFILECOUNTS;
  }

  public get description(): string {
    return 'Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityFileCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get the total number of files and how many of them were active across all
    group sites associated with an Microsoft 365 Group for the last week
      ${commands.O365GROUP_REPORT_ACTIVITYFILECOUNTS} --period D7

    Get the total number of files and how many of them were active across all
    group sites associated with an Microsoft 365 Group for the last week
    and exports the report data in the specified path in text format
      ${commands.O365GROUP_REPORT_ACTIVITYFILECOUNTS} --period D7 --output text > "o365groupactivityfilecounts.txt"

    Get the total number of files and how many of them were active across all
    group sites associated with an Microsoft 365 Group for the last week
    and exports the report data in the specified path in json format
      ${commands.O365GROUP_REPORT_ACTIVITYFILECOUNTS} --period D7 --output json > "o365groupactivityfilecounts.json"
`);
  }
}

module.exports = new O365GroupReportActivityFileCountsCommand();