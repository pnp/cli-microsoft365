import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportSiteUsagePagesCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_SITEUSAGEPAGES}`;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsagePages';
  }

  public get description(): string {
    return 'Gets the number of pages viewed across all sites';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of pages viewed across all sites for the last week
      ${commands.REPORT_SITEUSAGEPAGES} --period D7

    Gets the number of pages viewed across all sites for the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_SITEUSAGEPAGES} --period D7 --output text > "siteusagepages.txt"

    Gets the number of pages viewed across all sites for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_SITEUSAGEPAGES} --period D7 --output json > "siteusagepages.json"
`);
  }
}

module.exports = new SpoReportSiteUsagePagesCommand();