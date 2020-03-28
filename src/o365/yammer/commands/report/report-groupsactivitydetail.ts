import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class YammerReportGroupActivityDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityDetail';
  }

  public get description(): string {
    return 'Gets details about Yammer groups activity by group';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets details about Yammer groups activity by group for the last week
      ${commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL} --period D7

    Gets details about Yammer groups activity by group for May 1, 2019
      ${commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL} --date 2019-05-01

    Gets details about Yammer groups activity by group for the last week
    and exports the report data in the specified path in text format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL} --period D7 --output text --outputFile "groupsactivitydetail.txt"

    Gets details about Yammer groups activity by group for the last week
    and exports the report data in the specified path in json format
      ${commands.YAMMER_REPORT_GROUPSACTIVITYDETAIL} --period D7 --output json --outputFile "groupsactivitydetail.json"
`);
  }
}

module.exports = new YammerReportGroupActivityDetailCommand();