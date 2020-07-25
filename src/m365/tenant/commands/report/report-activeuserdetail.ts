import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class TenantReportActiveUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.TENANT_REPORT_ACTIVEUSERDETAIL;
  }

  public get description(): string {
    return 'Gets details about Microsoft 365 active users';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ActiveUserDetail';
  }
}

module.exports = new TenantReportActiveUserDetailCommand();