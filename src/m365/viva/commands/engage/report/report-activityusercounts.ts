import PeriodBasedReport from '../../../../base/PeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYUSERCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_ACTIVITYUSERCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on the number of unique users who posted, read, and liked Viva Engage messages';
  }
}

export default new YammerReportActivityUserCountsCommand();
