import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class OneDriveReportUsageFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_USAGEFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveUsageFileCounts';
  }

  public get description(): string {
    return 'Gets the total number of files across all sites and how many are active files';
  }
}

module.exports = new OneDriveReportUsageFileCountsCommand();