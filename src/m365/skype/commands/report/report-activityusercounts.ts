import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SkypeReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.SKYPE_REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on how many unique users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends on how many unique users organized and participated in
    conference  sessions held in your organization through Skype for Business.
    The report also includes the number of peer-to-peer sessions for the last
    week
      ${commands.SKYPE_REPORT_ACTIVITYUSERCOUNTS} --period D7

    Gets the trends on how many unique users organized and participated in
    conference sessions held in your organization through Skype for Business.
    The report also includes the number of peer-to-peer sessions for the last
    week and exports the report data in the specified path in text format
      ${commands.SKYPE_REPORT_ACTIVITYUSERCOUNTS} --period D7 --output text > "activityusercounts.txt"

    Gets the trends on how many unique users organized and participated in
    conference sessions held in your organization through Skype for Business.
    The report also includes the number of peer-to-peer sessions for the last
    week and exports the report data in the specified path in json format
      ${commands.SKYPE_REPORT_ACTIVITYUSERCOUNTS} --period D7 --output json > "activityusercounts.json"
`);
  }
}

module.exports = new SkypeReportActivityUserCountsCommand();
