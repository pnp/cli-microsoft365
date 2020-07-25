import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OneDriveReportUsageStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGESTORAGE;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageStorage';
  }

  public get description(): string {
    return 'Gets the trend on the amount of storage you are using in OneDrive for Business';
  }
}

module.exports = new OneDriveReportUsageStorageCommand();