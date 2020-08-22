import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class O365GroupReportActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return `${commands.O365GROUP_REPORT_ACTIVITYDETAIL}`;
  }

  public get description(): string {
    return 'Get details about Microsoft 365 Groups activity by group';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityDetail';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get details about Microsoft 365 Groups activity by group for  the last week
      ${commands.O365GROUP_REPORT_ACTIVITYDETAIL} --period D7

    Get details about Microsoft 365 Groups activity by group for September 28, 2019
      ${commands.O365GROUP_REPORT_ACTIVITYDETAIL} --date 2019-09-28

    Get details about Microsoft 365 Groups activity by group for the last week
    and exports the report data in the specified path in text format
      ${commands.O365GROUP_REPORT_ACTIVITYDETAIL} --period D7 --output text > "o365groupactivitydetail.txt"

    Get details about Microsoft 365 Groups activity by group for the last week
    and exports the report data in the specified path in json format
      ${commands.O365GROUP_REPORT_ACTIVITYDETAIL} --period D7 --output json > "o365groupactivitydetail.json"
`);
  }
}

module.exports = new O365GroupReportActivityDetailCommand();