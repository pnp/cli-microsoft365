import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OneDriveReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about OneDrive activity by user';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about OneDrive activity by user for the last week
      ${commands.REPORT_ACTIVITYUSERDETAIL} --period D7

    Gets details about OneDrive activity by user for May 1, 2019
      ${commands.REPORT_ACTIVITYUSERDETAIL} --date 2019-05-01

    Gets details about OneDrive activity by user for the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_ACTIVITYUSERDETAIL} --period D7 --output text > "onedriveactivityuserdetail.txt"

    Gets details about OneDrive activity by user for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_ACTIVITYUSERDETAIL} --period D7 --output json > "onedriveactivityuserdetail.json"
`);
  }
}

module.exports = new OneDriveReportActivityUserDetailCommand();