import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class O365GroupReportActivityFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYFILECOUNTS;
  }

  public get description(): string {
    return 'Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityFileCounts';
  }
}

export default new O365GroupReportActivityFileCountsCommand();