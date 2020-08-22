import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OneDriveReportUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGESTORAGE;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageStorage';
  }

  public get description(): string {
    return 'Gets the trend on the amount of storage you are using in OneDrive for Business';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trend on the amount of storage you are using in OneDrive for
    Business for the last week
      ${commands.REPORT_USAGESTORAGE} --period D7

    Gets the trend on the amount of storage you are using in OneDrive for
    Business for the last week and exports the report data in the specified path
    in text format
      ${commands.REPORT_USAGESTORAGE} --period D7 --output text > "usagestorage.txt"

    Gets the trend on the amount of storage you are using in OneDrive for
    Business for the last week and exports the report data in the specified path
    in json format
      ${commands.REPORT_USAGESTORAGE} --period D7 --output json > "usagestorage.json"
`);
  }
}

module.exports = new OneDriveReportUsageStorageCommand();