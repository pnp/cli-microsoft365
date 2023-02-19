import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class OneDriveReportUsageAccountDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGEACCOUNTDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountDetail';
  }

  public get description(): string {
    return 'Gets details about OneDrive usage by account';
  }
}

export default new OneDriveReportUsageAccountDetailCommand();