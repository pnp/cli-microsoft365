import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class SpoReportSiteUsageSiteCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_SITEUSAGESITECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageSiteCounts';
  }

  public get description(): string {
    return 'Gets the total number of files across all sites and the number of active files';
  }
}

export default new SpoReportSiteUsageSiteCountsCommand();