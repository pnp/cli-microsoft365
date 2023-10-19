import PeriodBasedReport from '../../../../base/PeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYCOUNTS;
  }

  public alias(): string[] {
    return [yammerCommands.REPORT_ACTIVITYCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on the amount of Viva Engage activity in your organization by how many messages were posted, read, and liked';
  }
}

export default new YammerReportActivityCountsCommand();
