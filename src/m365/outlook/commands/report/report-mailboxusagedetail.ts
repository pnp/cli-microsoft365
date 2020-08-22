import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailboxUsageDetailCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILBOXUSAGEDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getMailboxUsageDetail';
  }

  public get description(): string {
    return 'Gets details about mailbox usage';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about mailbox usage for the last week
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEDETAIL} --period D7

    Gets details about mailbox usage for the last week
    and exports the report data in the specified path in text format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEDETAIL} --period D7 --output text > "mailboxusagedetails.txt"

    Gets  details about mailbox usage for the last week
    and exports the report data in the specified path in json format
      ${commands.OUTLOOK_REPORT_MAILBOXUSAGEDETAIL} --period D7 --output json > "mailboxusagedetails.json"
`);
  }
}

module.exports = new OutlookReportMailboxUsageDetailCommand();