import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookGetMailboxUsageStroageCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_USAGE_GETMAILBOXUSAGESTORAGE}`;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageStorage';
  }

  public get description(): string {
    return 'Get the amount of storage used in your organization.';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the amount of storage used in your organization for the last week
      ${commands.OUTLOOK_USAGE_GETMAILBOXUSAGESTORAGE} --period D7

    Gets the amount of storage used in your organization for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_USAGE_GETMAILBOXUSAGESTORAGE} --period D7 --output text --outputFile 'C:/report.txt'

    Gets the amount of storage used in your organization for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_USAGE_GETMAILBOXUSAGESTORAGE} --period D7 --output json --outputFile 'C:/report.json'
`);
  }
}

module.exports = new OutlookGetMailboxUsageStroageCommand();