import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about email activity users have performed';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about email activity users have performed for the last week
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL} --period D7

    Gets details about email activity users have performed for May 1st, 2019
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL} --date 2019-05-01

    Gets details about email activity users have performed for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL} --period D7 --output text > "mailactivityuserdetails.txt"

    Gets details about email activity users have performed for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERDETAIL} --period D7 --output json > "mailactivityuserdetails.json"
`);
  }
}

module.exports = new OutlookReportMailActivityUserDetailCommand();