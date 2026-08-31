import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

export const options = globalOptionsZod.strict();

interface ManagementApp {
  applicationId: string
}

class PpManagementAppListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.MANAGEMENTAPP_LIST;
  }

  public get description(): string {
    return 'Lists management applications for Power Platform';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public async commandAction(logger: Logger): Promise<void> {
    const endpoint = `${this.resource}/providers/Microsoft.BusinessAppPlatform/adminApplications?api-version=2020-06-01`;

    try {
      const managementApps = await odata.getAllItems<ManagementApp>(endpoint);
      await logger.log(managementApps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpManagementAppListCommand();
