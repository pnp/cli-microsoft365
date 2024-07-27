import { Logger } from '../../../../cli/Logger.js';
import DateAndPeriodBasedReport, { CommandArgs } from '../../../base/DateAndPeriodBasedReport.js';
import commands from '../../commands.js';

class VivaEngageReportActivityUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.ENGAGE_REPORT_ACTIVITYUSERDETAIL;
  }

  public get usageEndpoint(): string {
    return 'getYammerActivityUserDetail';
  }

  public get description(): string {
    return 'Gets details about Viva Engage activity by user';
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await super.commandAction(logger, args);
  }
}

export default new VivaEngageReportActivityUserDetailCommand();