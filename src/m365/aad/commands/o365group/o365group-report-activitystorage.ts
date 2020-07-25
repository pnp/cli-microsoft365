import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';

class O365GroupReportActivityStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.O365GROUP_REPORT_ACTIVITYSTORAGE;
  }

  public get description(): string {
    return 'Get the total storage used across all group mailboxes and group sites';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityStorage';
  }
}

module.exports = new O365GroupReportActivityStorageCommand();