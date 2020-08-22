import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailboxUsageMailboxCountCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILBOXUSAGEMAILBOXCOUNT}`;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageMailboxCounts';
  }

  public get description(): string {
    return 'Gets the total number of user mailboxes in your organization and how many are active each day of the reporting period.';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
    
    A mailbox is considered active if the user sent or read any email.

  Examples:
    
    Gets the total number of user mailboxes in your organization and how many
    are active each day for the last week.
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEMAILBOXCOUNT} --period D7

    Gets the total number of user mailboxes in your organization and how many
    are active each day for the last week and exports the report data in the
    specified path in text format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEMAILBOXCOUNT} --period D7 --output text > "mailboxusagemailboxcounts.txt"

    Gets the total number of user mailboxes in your organization and how many
    are active each day for the last week and exports the report data in the
    specified path in json format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEMAILBOXCOUNT} --period D7 --output json > "mailboxusagemailboxcounts.json"
`);
  }
}

module.exports = new OutlookReportMailboxUsageMailboxCountCommand();
