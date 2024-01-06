import { Logger } from '../../../../cli/Logger.js';
import DateAndPeriodBasedReport, { CommandArgs } from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';
import yammerCommands from './yammerCommands.js';

class VivaEngageReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public alias(): string[] | undefined {
    return [yammerCommands.REPORT_DEVICEUSAGEUSERDETAIL];
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Viva Engage device usage by user';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportDeviceUsageUserDetailCommand();