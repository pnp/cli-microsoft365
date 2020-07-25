import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

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

module.exports = new O365GroupReportActivityFileCountsCommand();