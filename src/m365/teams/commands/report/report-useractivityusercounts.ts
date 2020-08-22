import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TeamsReportUserActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.';
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityUserCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

    Gets the number of Microsoft Teams users by activity type for last week
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7

    Gets the number of Microsoft Teams users by activity type for last week
    and exports the report data in the specified path in text format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7 --output text > "useractivityusercounts.txt"

    Gets the number of Microsoft Teams users by activity type for last week
    and exports the report data in the specified path in json format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERCOUNTS} --period D7 --output json > "useractivityusercounts.json"
`);
  }
}

module.exports = new TeamsReportUserActivityUserCountsCommand();