import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class TenantReportActiveUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TENANT_REPORT_ACTIVEUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Gets the count of daily active users in the reporting period by product.';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ActiveUserCounts';
  }
}

module.exports = new TenantReportActiveUserCountsCommand();