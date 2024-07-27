import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on the amount of VivaEngage activity in your organization by how many messages were posted, read, and liked';
  }
}

export default new VivaEngageReportActivityCountsCommand();
