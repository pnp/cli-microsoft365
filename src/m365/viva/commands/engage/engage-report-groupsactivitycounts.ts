import { Logger } from '../../../../cli/Logger.js';
import PeriodBasedReport, { CommandArgs } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

class VivaEngageReportGroupsActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_GROUPSACTIVITYCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_GROUPSACTIVITYCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerGroupsActivityCounts';
  }

  public get description(): string {
    return 'Gets the number of Viva Engage messages posted, read, and liked in groups';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportGroupsActivityCountsCommand();
