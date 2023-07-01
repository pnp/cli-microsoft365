import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class SpoReportSiteUsageFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_SITEUSAGEFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageFileCounts';
  }

  public get description(): string {
    return 'Get the total number of files across all sites and the number of active files';
  }
}

export default new SpoReportSiteUsageFileCountsCommand();