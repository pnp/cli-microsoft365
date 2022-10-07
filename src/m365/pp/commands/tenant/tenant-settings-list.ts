import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

class PpTenantSettingsListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the global Power Platform tenant settings';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const requestOptions: AxiosRequestConfig  = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/listtenantsettings?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };
    
    try {
      const res = await request.post<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpTenantSettingsListCommand();