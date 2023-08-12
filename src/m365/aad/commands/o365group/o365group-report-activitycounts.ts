import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class O365GroupReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYCOUNTS;
  }

  public get description(): string {
    return 'Get the number of group activities across group workloads';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityCounts';
  }
}

export default new O365GroupReportActivityCountsCommand();