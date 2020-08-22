import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoReportSiteUsageDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return `${commands.REPORT_SITEUSAGEDETAIL}`;
  }

  public get description(): string {
    return 'Gets details about SharePoint site usage';
  }

  public get usageEndpoint(): string {
    return 'getSharePointSiteUsageDetail';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    As this report is only available for the past 28 days, date parameter value
    should be a date from that range.

  Examples:

    Gets details about SharePoint site usage for the last week
      ${commands.REPORT_SITEUSAGEDETAIL} --period D7

    Gets details about SharePoint site usage for May 1, 2019
      ${commands.REPORT_SITEUSAGEDETAIL} --date 2019-05-01

    Gets details about SharePoint site usage for the last week 
    and exports the report data in the specified path in text format
      ${commands.REPORT_SITEUSAGEDETAIL} --period D7 --output text > "siteusagedetail.txt"

    Gets details about SharePoint site usage for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_SITEUSAGEDETAIL} --period D7 --output json > "siteusagedetail.json"
`);
  }
}

module.exports = new SpoReportSiteUsageDetailCommand();