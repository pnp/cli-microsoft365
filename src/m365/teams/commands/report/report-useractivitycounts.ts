import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TeamsReportUserActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

    Gets the number of Microsoft Teams activities by activity type for last week
      ${commands.TEAMS_REPORT_USERACTIVITYCOUNTS} --period D7

    Gets the number of Microsoft Teams activities by activity type for last week
    and exports the report data in the specified path in text format
      ${commands.TEAMS_REPORT_USERACTIVITYCOUNTS} --period D7 --output text > "useractivitycounts.txt"

    Gets the number of Microsoft Teams activities by activity type for last week
    and exports the report data in the specified path in json format
      ${commands.TEAMS_REPORT_USERACTIVITYCOUNTS} --period D7 --output json > "useractivitycounts.json"
`);
  }
}

module.exports = new TeamsReportUserActivityCountsCommand();