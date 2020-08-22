import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailAppUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users that connected to Exchange Online using any email app';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the count of unique users that connected to Exchange Online using any email app for the last week
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERCOUNTS} --period D7

    Gets the count of unique users that connected to Exchange Online using any email app for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERCOUNTS} --period D7 --output text > "mailappusageusercounts.txt"

    Gets the count of unique users that connected to Exchange Online using any email app for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERCOUNTS} --period D7 --output json > "mailappusageusercounts.json"
`);
  }
}

module.exports = new OutlookReportMailAppUsageUserCountsCommand();