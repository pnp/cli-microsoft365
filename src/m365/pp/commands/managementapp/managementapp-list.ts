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

  public commandAction(logger: Logger, args: any, cb: () => void): void {
    const endpoint = `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01`;

    odata
      .getAllItems<ManagementApp>(endpoint)
      .then((managementApps): void => {
        logger.log(managementApps);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new PpManagementAppListCommand();
