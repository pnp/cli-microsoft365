import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class YammerReportGroupsActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_GROUPSACTIVITYDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer groups activity by group';
  }
}

export default new YammerReportGroupsActivityDetailCommand();