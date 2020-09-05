import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OutlookReportMailAppUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.OUTLOOK_REPORT_MAILAPPUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about which activities users performed on the various email apps';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about which activities users performed on the various email apps for the last week
      m365 ${this.name} --period D7

    Gets details about which activities users performed on the various email apps for May 1st, 2019
      m365 ${this.name} --date 2019-05-01

    Gets details about which activities users performed on the various email apps for the last week
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "mailappusageuserdetails.txt"

    Gets details about which activities users performed on the various email apps for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "mailappusageuserdetails.json"
`);
  }
}

module.exports = new OutlookReportMailAppUsageUserDetailCommand();