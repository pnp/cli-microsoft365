import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

class M365GroupReportActivityFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.M365GROUP_REPORT_ACTIVITYFILECOUNTS;
  }

  public get description(): string {
    return 'Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_REPORT_ACTIVITYFILECOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityFileCounts';
  }
}

export default new M365GroupReportActivityFileCountsCommand();