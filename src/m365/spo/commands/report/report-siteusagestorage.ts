import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

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
}

module.exports = new SpoReportSiteUsageStorageCommand();