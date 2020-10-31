import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class SpoReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trend in the number of active users';
  }
}

module.exports = new SpoReportActivityUserCountsCommand();