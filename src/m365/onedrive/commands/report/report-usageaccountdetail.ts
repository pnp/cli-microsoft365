import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';
import commands from '../../commands';

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

module.exports = new OneDriveReportUsageAccountDetailCommand();