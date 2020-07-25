import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

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
}

module.exports = new SpoReportSiteUsageSiteCountsCommand();