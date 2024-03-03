
import { Logger } from '../../../../cli/Logger.js';
import PeriodBasedReport, { CommandArgs } from '../../../base/PeriodBasedReport.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

class VivaEngageReportDeviceUsageUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERCOUNTS;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEUSERCOUNTS];
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserCounts';
  }

  public get description(): string {
    return 'Gets the number of daily users by device type';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportDeviceUsageUserCountsCommand();
