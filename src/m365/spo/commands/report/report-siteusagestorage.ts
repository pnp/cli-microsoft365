import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportSiteUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_SITEUSAGESTORAGE}`;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageStorage';
  }

  public get description(): string {
    return 'Gets the trend of storage allocated and consumed during the reporting period';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trend of storage allocated and consumed during the last week
      ${commands.REPORT_SITEUSAGESTORAGE} --period D7

    Gets the trend of storage allocated and consumed during the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_SITEUSAGESTORAGE} --period D7 --output text > "siteusagestorage.txt"

    Gets the trend of storage allocated and consumed during the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_SITEUSAGESTORAGE} --period D7 --output json > "siteusagestorage.json"
`);
  }
}

module.exports = new SpoReportSiteUsageStorageCommand();