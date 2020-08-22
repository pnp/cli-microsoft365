import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SkypeReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.SKYPE_REPORT_ACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends on how many users organized and participated in conference 
    sessions held in your organization through Skype for Business. The report 
    also includes the number of peer-to-peer sessions for the last week
      ${commands.SKYPE_REPORT_ACTIVITYCOUNTS} --period D7

    Gets the trends on how many users organized and participated in conference 
    sessions held in your organization through Skype for Business. The report 
    also includes the number of peer-to-peer sessions for the last week and 
    exports the report data in the specified path in text format
      ${commands.SKYPE_REPORT_ACTIVITYCOUNTS} --period D7 --output text > "activitycounts.txt"

    Gets the trends on how many users organized and participated in conference 
    sessions held in your organization through Skype for Business. The report 
    also includes the number of peer-to-peer sessions for the last week and 
    exports the report data in the specified path in json format
      ${commands.SKYPE_REPORT_ACTIVITYCOUNTS} --period D7 --output json > "activitycounts.json"
`);
  }
}

module.exports = new SkypeReportActivityCountsCommand();
