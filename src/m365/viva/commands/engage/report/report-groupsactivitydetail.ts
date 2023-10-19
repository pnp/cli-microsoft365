import DateAndPeriodBasedReport from '../../../../base/DateAndPeriodBasedReport.js';
import commands from '../../../../viva/commands.js';
import yammerCommands from '../../../../yammer/commands.js';

class YammerReportGroupsActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_GROUPSACTIVITYDETAIL;
  }

  public alias(): string[] {
    return [yammerCommands.REPORT_GROUPSACTIVITYDETAIL];
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer groups activity by group';
  }
}

export default new YammerReportGroupsActivityDetailCommand();