import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportOneDriveUsageAccountCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ONEDRIVEUSAGEACCOUNTCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active OneDrive for Business sites';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    Any site on which users viewed, modified, uploaded, downloaded, shared, or synced files is considered an active site.
      
         Examples:
      
    Gets the trend in the number of active OneDrive for Business sites for the last week
      ${commands.REPORT_ONEDRIVEUSAGEACCOUNTCOUNTS} --period D7

    Gets the trend in the number of active OneDrive for Business sites for the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_ONEDRIVEUSAGEACCOUNTCOUNTS} --period D7 --output text --outputFile 'C:/report.txt'

    Gets the trend in the number of active OneDrive for Business sites for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_ONEDRIVEUSAGEACCOUNTCOUNTS} --period D7 --output json --outputFile 'C:/report.json'
`);
  }
}

module.exports = new SpoReportOneDriveUsageAccountCountsCommand();