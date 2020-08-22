import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of unique pages visited by users for the last week
      ${commands.REPORT_ACTIVITYPAGES} --period D7

    Gets the number of unique pages visited by users for the last week
    and exports the report data in the specified path in text format
      ${commands.REPORT_ACTIVITYPAGES} --period D7 --output text > "activitypages.txt"

    Gets the number of unique pages visited by users for the last week
    and exports the report data in the specified path in json format
      ${commands.REPORT_ACTIVITYPAGES} --period D7 --output json > "activitypages.json"
`);
  }
}

module.exports = new SpoReportActivityPagesCommand();