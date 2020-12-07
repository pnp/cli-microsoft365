import TeamsCallReport from '../../../base/TeamsCallReport';
import commands from '../../commands';

class TeamsReportPstnCallsCommand extends TeamsCallReport {
  public get name(): string {
    return `${commands.TEAMS_REPORT_PSTNCALLS}`;
  }

  public get description(): string {
    return 'Get details about PSTN calls made within a given time period';
  }

  public get usageEndpoint(): string {
    return 'getPstnCalls';
  }
}

module.exports = new TeamsReportPstnCallsCommand();