import { Logger } from '../../../../cli/Logger.js';
import PeriodBasedReport, { CommandArgs } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

class VivaEngageReportActivityUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYUSERCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_ACTIVITYUSERCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserCounts';
  }

  public get description(): string {
    return 'Gets the trends on the number of unique users who posted, read, and liked Viva Engage messages';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportActivityUserCountsCommand();
