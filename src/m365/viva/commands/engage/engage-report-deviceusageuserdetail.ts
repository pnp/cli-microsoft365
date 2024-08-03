import { Logger } from '../../../../cli/Logger.js';
import DateAndPeriodBasedReport, { CommandArgs } from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportDeviceUsageUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_DEVICEUSAGEUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerDeviceUsageUserDetail';
  }

  public get description(): string {
    return 'Gets details about Viva Engage device usage by user';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportDeviceUsageUserDetailCommand();