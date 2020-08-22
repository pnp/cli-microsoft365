import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailboxUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILBOXUSAGESTORAGE}`;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageStorage';
  }

  public get description(): string {
    return 'Gets the amount of mailbox storage used in your organization';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the amount of mailbox storage used in your organization for the last week
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGESTORAGE} --period D7

    Gets the amount of mailbox storage used in your organization for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGESTORAGE} --period D7 --output text > "mailboxusagestorage.txt"

    Gets the amount of mailbox storage used in your organization for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGESTORAGE} --period D7 --output json > "mailboxusagestorage.json"
`);
  }
}

module.exports = new OutlookReportMailboxUsageStorageCommand();