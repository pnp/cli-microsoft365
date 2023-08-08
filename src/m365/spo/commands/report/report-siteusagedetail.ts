import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class SpoReportSiteUsageDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_SITEUSAGEDETAIL;
  }

  public get description(): string {
    return 'Gets details about SharePoint site usage';
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageDetail';
  }
}

export default new SpoReportSiteUsageDetailCommand();