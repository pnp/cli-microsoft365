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

  public async commandAction(logger: Logger): Promise<void> {
    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const result = await request.get(requestOptions);
      logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PlannerTenantSettingsListCommand();