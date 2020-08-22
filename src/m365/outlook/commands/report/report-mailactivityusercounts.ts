import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILACTIVITYUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityUserCounts';
  }

  public get description(): string {
    return 'Enables you to understand trends on the number of unique users who are performing email activities like send, read, and receive';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERCOUNTS} --period D7

    Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERCOUNTS} --period D7 --output text > "mailactivityusercounts.txt"

    Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILACTIVITYUSERCOUNTS} --period D7 --output json > "mailactivityusercounts.json"
`);
  }
}

module.exports = new OutlookReportMailActivityUserCountsCommand();