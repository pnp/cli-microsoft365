import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class O365GroupReportActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYDETAIL;
  }

  public get description(): string {
    return 'Get details about Microsoft 365 Groups activity by group';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityDetail';
  }
}

export default new O365GroupReportActivityDetailCommand();