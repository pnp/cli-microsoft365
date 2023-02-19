import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

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

export default new OneDriveReportUsageFileCountsCommand();