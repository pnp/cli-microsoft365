import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class SpoReportActivityFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityFileCounts';
  }

  public get description(): string {
    return 'Gets the number of unique, licensed users who interacted with files stored on SharePoint sites';
  }
}

export default new SpoReportActivityFileCountsCommand();
