import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class YammerReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Yammer messages posted, read, and liked in groups';
  }
}

module.exports = new YammerReportGroupsActivityCountsCommand();
