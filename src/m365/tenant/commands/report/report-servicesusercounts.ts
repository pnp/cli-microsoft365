import PeriodBasedReport, { periodBasedReportOptions as options } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export { options };

class TenantReportServicesUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_SERVICESUSERCOUNTS;
  }

  public get description(): string {
    return 'Gets the count of users by activity type and service.';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ServicesUserCounts';
  }
}

export default new TenantReportServicesUserCountsCommand();