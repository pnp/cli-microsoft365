import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class TenantReportServicesUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.TENANT_REPORT_SERVICESUSERCOUNTS;
  }

  public get description(): string {
    return 'Gets the count of users by activity type and service.';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ServicesUserCounts';
  }
}

module.exports = new TenantReportServicesUserCountsCommand();