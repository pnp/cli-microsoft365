import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import auth from '../../../../Auth.js';
import { Application, RequiredResourceAccess } from '@microsoft/microsoft-graph-types';
import config from '../../../../config.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
import { settingsNames } from '../../../../settingsNames.js';
import { browserUtil } from '../../../../utils/browserUtil.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';

class CliAppReconsentCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_RECONSENT;
  }

  public get description(): string {
    return 'Reconsent all permission scopes used in CLI for Microsoft 365';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const appId = auth.connection.appId!;
      if (this.verbose) {
        await logger.logToStderr(`Adding all missing permission scopes used in CLI for Microsoft 365 to application with ID '${appId}'...`);
      }

      const application = await entraApp.getAppRegistrationByAppId(appId, ['requiredResourceAccess', 'id']);
      await this.addCliAppScopes(logger, application.requiredResourceAccess!);
      await this.updateAppScopes(logger, application.id!, application.requiredResourceAccess!);

      const consentUrl = `https://login.microsoftonline.com/${auth.connection.tenant}/adminconsent?client_id=${appId}`;
      await logger.log(`To consent to the new scopes for your Microsoft Entra application registration, please navigate to the following URL: ${consentUrl}`);

      if (cli.getSettingWithDefaultValue(settingsNames.autoOpenLinksInBrowser, false)) {
        await browserUtil.open(consentUrl);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addCliAppScopes(logger: Logger, appScopes: RequiredResourceAccess[]): Promise<void> {
    const allCliScopes = config.allScopes;
    const servicePrincipals = await entraServicePrincipal.getServicePrincipals('displayName,appId,oauth2PermissionScopes,servicePrincipalNames');

    if (this.verbose) {
      await logger.logToStderr(`Verifying if all ${allCliScopes.length} permission scopes are present in the app registration...`);
    }

    for (const cliScope of allCliScopes) {
      // Extract service principal name and scope from the URL string
      const spName = urlUtil.removeTrailingSlashes(cliScope.substring(0, cliScope.lastIndexOf('/')));
      const scopeName = cliScope.substring(cliScope.lastIndexOf('/') + 1);

      // Find the matching service principal by name
      const servicePrincipal = servicePrincipals.find(sp => sp.servicePrincipalNames?.some(name => urlUtil.removeTrailingSlashes(name).toLowerCase() === spName.toLowerCase()));

      if (!servicePrincipal) {
        if (this.verbose) {
          await logger.logToStderr(`Service principal with name '${spName}' not found. Skipping scope '${scopeName}'.`);
        }
        continue;
      }

      // Find the matching scope in the service principal
      const scope = servicePrincipal.oauth2PermissionScopes?.find(s => s.value?.toLowerCase() === scopeName.toLowerCase());

      if (!scope) {
        if (this.verbose) {
          await logger.logToStderr(`Scope '${scopeName}' not found in service principal '${spName}'. Skipping scope...`);
        }
        continue;
      }

      // Check if the service principal is already present in the app registration
      let appSp = appScopes.find(sp => sp.resourceAppId?.toLowerCase() === servicePrincipal.appId!.toLowerCase());
      if (!appSp) {
        // Service principal is not present in the app registration, let's add it
        appSp = {
          resourceAppId: servicePrincipal.appId!,
          resourceAccess: []
        };
        appScopes.push(appSp);
      }

      // Check if the scope is already present in the app registration
      const isAppScopePresent = appSp.resourceAccess!.some(s => s.id?.toLowerCase() === scope.id!.toLowerCase());
      if (!isAppScopePresent) {
        // Scope is not present in the app registration, let's add it
        appSp.resourceAccess!.push({
          id: scope.id!,
          type: 'Scope'
        });
      }
    }
  }

  private async updateAppScopes(logger: Logger, appId: string, appScopes: RequiredResourceAccess[]): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating permission scopes of application with ID '${appId}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/applications/${appId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        requiredResourceAccess: appScopes
      } as Application
    };

    await request.patch(requestOptions);
  }
}

export default new CliAppReconsentCommand();