import PeriodBasedReport from '../../../../base/PeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS;
  }

  public alias(): string[] {
    return [yammerCommands.REPORT_GROUPSACTIVITYCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Yammer messages posted, read, and liked in groups';
  }
}

export default new YammerReportGroupsActivityCountsCommand();
