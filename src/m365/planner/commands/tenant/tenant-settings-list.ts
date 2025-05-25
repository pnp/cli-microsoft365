import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PlannerCommand from '../../../base/PlannerCommand.js';
import commands from '../../commands.js';

class PlannerTenantSettingsListCommand extends PlannerCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the Microsoft Planner configuration of the tenant';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/taskAPI/tenantAdminSettings/Settings`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const result = await request.get(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PlannerTenantSettingsListCommand();