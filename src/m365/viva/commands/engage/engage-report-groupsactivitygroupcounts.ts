import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportGroupsActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_GROUPSACTIVITYGROUPCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityGroupCounts';
  }

  public get description(): string {
    return 'Gets the total number of groups that existed and how many included group conversation activity';
  }
}

export default new VivaEngageReportGroupsActivityGroupCountsCommand();