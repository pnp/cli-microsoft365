import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class O365GroupReportActivityStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYSTORAGE;
  }

  public get description(): string {
    return 'Get the total storage used across all group mailboxes and group sites';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityStorage';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get the total storage used across all group mailboxes and group sites for
    the last week
      ${commands.O365GROUP_REPORT_ACTIVITYSTORAGE} --period D7

    Get the total storage used across all group mailboxes and group sites and
    exports the report data in the specified path in text format
      ${commands.O365GROUP_REPORT_ACTIVITYSTORAGE} --period D7 --output text > "o365groupactivitystorage.txt"

    Get the total storage used across all group mailboxes and group sites for
    the last week and exports the report data in the specified path in json
    format
      ${commands.O365GROUP_REPORT_ACTIVITYSTORAGE} --period D7 --output json > "o365groupactivitystorage.json"
`);
  }
}

module.exports = new O365GroupReportActivityStorageCommand();