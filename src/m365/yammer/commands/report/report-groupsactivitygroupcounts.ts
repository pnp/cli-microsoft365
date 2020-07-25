import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class YammerReportGroupsActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYGROUPCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityGroupCounts';
  }

  public get description(): string {
    return 'Gets the total number of groups that existed and how many included group conversation activity';
  }
}

module.exports = new YammerReportGroupsActivityGroupCountsCommand();