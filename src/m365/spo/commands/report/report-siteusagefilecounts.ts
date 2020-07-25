import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class SpoReportSiteUsageFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_SITEUSAGEFILECOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageFileCounts';
  }

  public get description(): string {
    return 'Get the total number of files across all sites and the number of active files';
  }
}

module.exports = new SpoReportSiteUsageFileCountsCommand();