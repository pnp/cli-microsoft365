import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailAppUsageAppsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILAPPUSAGEAPPSUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageAppsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users per email app';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the count of unique users per email app for the last week
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEAPPSUSERCOUNTS} --period D7

    Gets the count of unique users per email app for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEAPPSUSERCOUNTS} --period D7 --output text > "mailappusageappusercounts.txt"

    Gets the count of unique users per email app for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILAPPUSAGEAPPSUSERCOUNTS} --period D7 --output json > "mailappusageappusercounts.json"
`);
  }
}

module.exports = new OutlookReportMailAppUsageAppsUserCountsCommand();