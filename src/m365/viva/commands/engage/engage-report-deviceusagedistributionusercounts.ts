import PeriodBasedReport, { periodBasedReportOptions } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';

export const options = periodBasedReportOptions;

class VivaEngageReportDeviceUsageDistributionUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEDISTRIBUTIONUSERCOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageDistributionUserCounts';
  }

  public get description(): string {
    return 'Gets the number of users by device type';
  }
}

export default new VivaEngageReportDeviceUsageDistributionUserCountsCommand();

