import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class YammerReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on the number of unique users who posted, read, and liked Yammer messages';
  }
}

module.exports = new YammerReportActivityUserCountsCommand();
