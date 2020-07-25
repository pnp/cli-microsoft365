import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class OneDriveReportUsageAccountCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGEACCOUNTCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageAccountCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active OneDrive for Business sites';
  }
}

module.exports = new OneDriveReportUsageAccountCountsCommand();