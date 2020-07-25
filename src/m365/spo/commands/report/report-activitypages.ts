import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class SpoReportActivityPagesCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYPAGES;
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityPages';
  }

  public get description(): string {
    return 'Gets the number of unique pages visited by users';
  }
}

module.exports = new SpoReportActivityPagesCommand();