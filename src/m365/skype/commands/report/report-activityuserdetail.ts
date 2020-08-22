import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SkypeReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.SKYPE_REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getSkypeForBusinessActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Skype for Business activity by user';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about Skype for Business activity by user for the last week
      ${commands.SKYPE_REPORT_ACTIVITYUSERDETAIL} --period D7

    Gets details about Skype for Business activity by user for May 1, 2019
      ${commands.SKYPE_REPORT_ACTIVITYUSERDETAIL} --date 2019-05-01

    Gets details about Skype for Business activity by user for the last week
    and exports the report data in the specified path in text format
      ${commands.SKYPE_REPORT_ACTIVITYUSERDETAIL} --period D7 --output text > "activityuserdetail.txt"

    Gets details about Skype for Business activity by user for the last week
    and exports the report data in the specified path in json format
      ${commands.SKYPE_REPORT_ACTIVITYUSERDETAIL} --period D7 --output json > "activityuserdetail.json"
`);
  }
}

module.exports = new SkypeReportActivityUserDetailCommand();
