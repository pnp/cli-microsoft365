import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import PowerPlatformCommand from '../../../base/PowerPlatformCommand';
import commands from '../../commands';

class PpTenantSettingsListCommand extends PowerPlatformCommand {
  public get name(): string {
    return commands.TENANT_SETTINGS_LIST;
  }

  public get description(): string {
    return 'Lists the global Power Platform tenant settings';
  }

  public defaultProperties(): string[] | undefined {
    return ['disableCapacityAllocationByEnvironmentAdmins', 'disableEnvironmentCreationByNonAdminUsers', 'disableNPSCommentsReachout', 'disablePortalsCreationByNonAdminUsers', 'disableSupportTicketsVisibleByAllUsers', 'disableSurveyFeedback', 'disableTrialEnvironmentCreationByNonAdminUsers', 'walkMeOptOut'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    const requestOptions: CliRequestOptions = {
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