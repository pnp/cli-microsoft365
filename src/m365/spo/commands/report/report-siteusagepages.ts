import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class SpoReportSiteUsagePagesCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_SITEUSAGEPAGES;
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsagePages';
  }

  public get description(): string {
    return 'Gets the number of pages viewed across all sites';
  }
}

module.exports = new SpoReportSiteUsagePagesCommand();