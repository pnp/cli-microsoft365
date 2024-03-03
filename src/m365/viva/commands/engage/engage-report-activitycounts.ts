import { Logger } from '../../../../cli/Logger.js';
import PeriodBasedReport, { CommandArgs } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

class VivaEngageReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_ACTIVITYCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityCounts';
  }

  public get description(): string {
    return 'Gets the trends on the amount of VivaEngage activity in your organization by how many messages were posted, read, and liked';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportActivityCountsCommand();
