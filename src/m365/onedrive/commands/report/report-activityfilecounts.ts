import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class OneDriveReportActivityFileCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityFileCounts';
  }

  public get description(): string {
    return 'Gets the number of unique, licensed users that performed file interactions against any OneDrive account';
  }
}

export default new OneDriveReportActivityFileCountCommand();