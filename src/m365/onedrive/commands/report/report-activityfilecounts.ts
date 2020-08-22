import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

class OneDriveReportActivityFileCountCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getOneDriveActivityFileCounts';
  }

  public get description(): string {
    return 'Gets the number of unique, licensed users that performed file interactions against any OneDrive account';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of unique, licensed users that performed file interactions
    against any OneDrive account for the last week
      ${commands.REPORT_ACTIVITYFILECOUNTS} --period D7

    Gets the number of unique, licensed users that performed file interactions
    against any OneDrive account for the last week and exports the report data
    in the specified path in text format
      ${commands.REPORT_ACTIVITYFILECOUNTS} --period D7 --output text > "activityfilecounts.txt"

    Gets the number of unique, licensed users that performed file interactions
    against any OneDrive account for the last week and exports the report data
    in the specified path in json format
      ${commands.REPORT_ACTIVITYFILECOUNTS} --period D7 --output json > "activityfilecounts.json"
`);
  }
}

module.exports = new OneDriveReportActivityFileCountCommand();