import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport.js';
import aadCommands from '../../aadCommands.js';
import commands from '../../commands.js';

class M365GroupReportActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.M365GROUP_REPORT_ACTIVITYDETAIL;
  }

  public get description(): string {
    return 'Get details about Microsoft 365 Groups activity by group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_REPORT_ACTIVITYDETAIL];
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityDetail';
  }
}

export default new M365GroupReportActivityDetailCommand();