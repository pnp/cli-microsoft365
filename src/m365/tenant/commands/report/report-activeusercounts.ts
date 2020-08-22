import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class TenantReportActiveUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.TENANT_REPORT_ACTIVEUSERCOUNTS}`;
  }

  public get description(): string {
    return 'Gets the count of daily active users in the reporting period by product.';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ActiveUserCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

    Gets the count of daily active users in the reporting period by product for last week
      ${commands.TENANT_REPORT_ACTIVEUSERCOUNTS} --period D7

    Gets the count of daily active users in the reporting period by product for last week
    and exports the report data in the specified path in text format
      ${commands.TENANT_REPORT_ACTIVEUSERCOUNTS} --period D7 --output text > "activeusercounts.txt"

    Gets the count of daily active users in the reporting period by product for last week
    and exports the report data in the specified path in json format
      ${commands.TENANT_REPORT_ACTIVEUSERCOUNTS} --period D7 --output json > "activeusercounts.json"
`);
  }
}

module.exports = new TenantReportActiveUserCountsCommand();