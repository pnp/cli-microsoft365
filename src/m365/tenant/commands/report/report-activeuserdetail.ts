import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class TenantReportActiveUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVEUSERDETAIL;
  }

  public get description(): string {
    return 'Gets details about Microsoft 365 active users';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ActiveUserDetail';
  }
}

export default new TenantReportActiveUserDetailCommand();