import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportSiteUsageSiteCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_SITEUSAGESITECOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageSiteCounts';
  }

  public get description(): string {
    return 'Gets the total number of files across all sites and the number of active files';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    A file (user or system) is considered active if it has been saved, synced,
    modified, or shared within the specified time period.

  Examples:
      
    Gets the total number of files across all sites and the number of active
    files for the last week
      ${commands.REPORT_SITEUSAGESITECOUNTS} --period D7

    Gets the total number of files across all sites and the number of active
    files for the last week and exports the report data in the specified path in
    text format
      ${commands.REPORT_SITEUSAGESITECOUNTS} --period D7 --output text > "siteusagesitecounts.txt"

    Gets the total number of files across all sites and the number of active
    files for the last week and exports the report data in the specified path in
    json format
      ${commands.REPORT_SITEUSAGESITECOUNTS} --period D7 --output json > "siteusagesitecounts.json"
`);
  }
}

module.exports = new SpoReportSiteUsageSiteCountsCommand();