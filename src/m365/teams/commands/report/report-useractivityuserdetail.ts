import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TeamsReportUserActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL}`;
  }

  public get usageEndpoint(): string {
    return 'getTeamsUserActivityUserDetail';
  }

  public get description(): string {
    return 'Get details about Microsoft Teams user activity by user.';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Gets details about Microsoft Teams user activity by user for 
    the last week
      ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period D7

    Gets details about Microsoft Teams user activity by user 
    for July 13, 2019
      ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --date 2019-07-13

    Gets details about Microsoft Teams user activity by user for the last week
    and exports the report data in the specified path in text format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period D7 --output text > "useractivityuserdetails.txt"

    Gets details about Microsoft Teams user activity by user for the last week
    and exports the report data in the specified path in json format
      ${commands.TEAMS_REPORT_USERACTIVITYUSERDETAIL} --period D7 --output json > "useractivityuserdetails.json"
`);
  }
}

module.exports = new TeamsReportUserActivityUserDetailCommand();