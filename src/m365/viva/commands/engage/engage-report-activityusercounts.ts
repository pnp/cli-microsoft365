import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on the number of unique users who posted, read, and liked Viva Engage messages';
  }
}

export default new VivaEngageReportActivityUserCountsCommand();
