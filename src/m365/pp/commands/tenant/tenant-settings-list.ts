import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod.strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PpTenantSettingsListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the global Power Platform tenant settings';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public defaultProperties(): string[] | undefined {
    return ['disableCapacityAllocationByEnvironmentAdmins', 'disableEnvironmentCreationByNonAdminUsers', 'disableNPSCommentsReachout', 'disablePortalsCreationByNonAdminUsers', 'disableSupportTicketsVisibleByAllUsers', 'disableSurveyFeedback', 'disableTrialEnvironmentCreationByNonAdminUsers', 'walkMeOptOut'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/providers/Microsoft.BusinessAppPlatform/listtenantsettings?api-version=2020-10-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      if (args) {
        const res = await request.post<any>(requestOptions);
        await logger.log(res);
      } else {
        throw 'Invalid arguments';
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PpTenantSettingsListCommand();