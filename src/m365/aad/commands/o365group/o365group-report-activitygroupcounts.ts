import PeriodBasedReport from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

class O365GroupReportActivityGroupCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYGROUPCOUNTS;
  }

  public get description(): string {
    return 'Get the daily total number of groups and how many of them were active based on email conversations, Yammer posts, and SharePoint file activities';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityGroupCounts';
  }
}

export default new O365GroupReportActivityGroupCountsCommand();