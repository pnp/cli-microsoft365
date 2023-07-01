import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class YammerReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_GROUPSACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Yammer messages posted, read, and liked in groups';
  }
}

export default new YammerReportGroupsActivityCountsCommand();
