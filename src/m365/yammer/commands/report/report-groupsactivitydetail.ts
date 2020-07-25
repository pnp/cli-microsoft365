import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

class YammerReportGroupsActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer groups activity by group';
  }
}

module.exports = new YammerReportGroupsActivityDetailCommand();