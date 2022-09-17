import { Logger } from '../../../../cli';
import { odata } from '../../../../utils';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

export interface ManagementApp {
  applicationId: string
}

class PpManagementAppListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.MANAGEMENTAPP_LIST;
  }

  public get description(): string {
    return 'Lists management applications for Power Platform';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const endpoint = `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01`;

    try {
      const managementApps = await odata.getAllItems<ManagementApp>(endpoint);
      logger.log(managementApps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PpManagementAppListCommand();
