import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageQuotaStatusMailboxCounts';
  }

  public get description(): string {
    return 'Gets the count of user mailboxes in each quota category';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the count of user mailboxes in each quota category for the last week
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS} --period D7

    Gets the count of user mailboxes in each quota category for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS} --period D7 --output text > "mailboxusagequotastatusmailboxcounts.txt"

    Gets the count of user mailboxes in each quota category for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEQUOTASTATUSMAILBOXCOUNTS} --period D7 --output json > "mailboxusagequotastatusmailboxcounts.json"
`);
  }
}

module.exports = new OutlookReportMailboxUsageQuotaStatusMailboxCountsCommand();