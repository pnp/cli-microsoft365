import { Logger } from '../../../../cli';
import { AxiosRequestConfig } from 'axios';
import request from '../../../../request';
import PlannerCommand from '../../../base/PlannerCommand';
import commands from '../../commands';

class PlannerTenantSettingsListCommand extends PlannerCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the Microsoft Planner configuration of the tenant';
  }

  public defaultProperties(): string[] | undefined {
    return ['isPlannerAllowed', 'allowCalendarSharing', 'allowTenantMoveWithDataLoss', 'allowTenantMoveWithDataMigration', 'allowRosterCreation', 'allowPlannerMobilePushNotifications'];
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((result): void => {
        logger.log(result);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new PlannerTenantSettingsListCommand();