import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Viva Engage messages posted, read, and liked in groups';
  }
}

export default new VivaEngageReportGroupsActivityCountsCommand();
