import { v4 } from 'uuid';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';

interface ServicePrincipalInfo {
  appId: string;
  appRoles: { id: string; value: string; }[];
  oauth2PermissionScopes: { id: string; value: string; }[];
  servicePrincipalNames: string[];
}

interface RequiredResourceAccess {
  resourceAppId: string;
  resourceAccess: ResourceAccess[];
}

interface ResourceAccess {
  id: string;
  type: string;
}

interface AppInfo {
  appId: string;
  // objectId
  id: string;
  secret?: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  apisApplication?: string;
  apisDelegated?: string;
  implicitFlow: boolean;
  multitenant: boolean;
  name: string;
  platform?: string;
  redirectUris?: string;
  scopeAdminConsentDescription?: string;
  scopeAdminConsentDisplayName?: string;
  scopeConsentBy?: string;
  scopeName?: string;
  uri?: string;
  withSecret: boolean;
}

class AadAppAddCommand extends GraphItemsListCommand<ServicePrincipalInfo> {
  private static aadApplicationPlatform: string[] = ['spa', 'web', 'publicClient'];
  private static aadAppScopeConsentBy: string[] = ['admins', 'adminsAndUsers'];

  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Creates new Azure AD app registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.apis = typeof args.options.apisDelegated !== 'undefined';
    telemetryProps.implicitFlow = args.options.implicitFlow;
    telemetryProps.multitenant = args.options.multitenant;
    telemetryProps.platform = args.options.platform;
    telemetryProps.redirectUris = typeof args.options.redirectUris !== 'undefined';
    telemetryProps.scopeAdminConsentDescription = typeof args.options.scopeAdminConsentDescription !== 'undefined';
    telemetryProps.scopeAdminConsentDisplayName = typeof args.options.scopeAdminConsentDisplayName !== 'undefined';
    telemetryProps.scopeName = args.options.scopeConsentBy;
    telemetryProps.scopeName = typeof args.options.scopeName !== 'undefined';
    telemetryProps.uri = typeof args.options.uri !== 'undefined';
    telemetryProps.withSecret = args.options.withSecret;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .resolveApis(args, logger)
      .then(apis => this.createAppRegistration(args, apis, logger))
      .then(appInfo => this.configureUri(args, appInfo, logger))
      .then(appInfo => this.configureSecret(args, appInfo, logger))
      .then((_appInfo: AppInfo): void => {
        const appInfo: any = {
          appId: _appInfo.appId,
          objectId: _appInfo.id
        };
        if (_appInfo.secret) {
          appInfo.secret = _appInfo.secret;
        }

        logger.log(appInfo);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  private createAppRegistration(args: CommandArgs, apis: RequiredResourceAccess[], logger: Logger): Promise<AppInfo> {
    const applicationInfo: any = {
      displayName: args.options.name,
      signInAudience: args.options.multitenant ? 'AzureADMultipleOrgs' : 'AzureADMyOrg'
    };

    if (apis.length > 0) {
      applicationInfo.requiredResourceAccess = apis;
    }

    if (args.options.redirectUris) {
      applicationInfo[args.options.platform!] = {
        redirectUris: args.options.redirectUris.split(',').map(u => u.trim())
      };
    }

    if (args.options.implicitFlow) {
      if (!applicationInfo.web) {
        applicationInfo.web = {};
      }
      applicationInfo.web.implicitGrantSettings = {
        enableAccessTokenIssuance: true,
        enableIdTokenIssuance: true
      };
    }

    if (this.verbose) {
      logger.logToStderr(`Creating Azure AD app registration...`);
    }

    const createApplicationRequestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request.post<AppInfo>(createApplicationRequestOptions);
  }

  private configureUri(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.uri) {
      return Promise.resolve(appInfo);
    }

    if (this.verbose) {
      logger.logToStderr(`Configuring Azure AD application ID URI...`);
    }

    const applicationInfo: any = {};

    if (args.options.uri) {
      const appUri: string = args.options.uri.replace(/_appId_/g, appInfo.appId);
      applicationInfo.identifierUris = [appUri];
    }

    if (args.options.scopeName) {
      applicationInfo.api = {
        oauth2PermissionScopes: [{
          adminConsentDescription: args.options.scopeAdminConsentDescription,
          adminConsentDisplayName: args.options.scopeAdminConsentDisplayName,
          id: v4(),
          type: args.options.scopeConsentBy === 'adminsAndUsers' ? 'User' : 'Admin',
          value: args.options.scopeName
        }]
      };
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    return request
      .patch(requestOptions)
      .then(_ => appInfo);
  }

  private resolveApis(args: CommandArgs, logger: Logger): Promise<RequiredResourceAccess[]> {
    if (!args.options.apisDelegated && !args.options.apisApplication) {
      return Promise.resolve([]);
    }

    if (this.verbose) {
      logger.logToStderr('Resolving requested APIs...');
    }

    return this
      .getAllItems(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles`, logger, true)
      .then(_ => {
        try {
          const resolvedApis = this.getRequiredResourceAccessForApis(args.options.apisDelegated, 'Scope', logger);
          if (this.debug) {
            logger.logToStderr(`Resolved delegated permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
          }
          const resolvedApplicationApis = this.getRequiredResourceAccessForApis(args.options.apisApplication, 'Role', logger);
          if (this.debug) {
            logger.logToStderr(`Resolved application permissions: ${JSON.stringify(resolvedApplicationApis, null, 2)}`);
          }
          // merge resolved application APIs onto resolved delegated APIs
          resolvedApplicationApis.forEach(resolvedRequiredResource => {
            let requiredResource = resolvedApis.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
            if (requiredResource) {
              requiredResource.resourceAccess.push(...resolvedRequiredResource.resourceAccess);
            }
            else {
              resolvedApis.push(resolvedRequiredResource);
            }
          });

          if (this.debug) {
            logger.logToStderr(`Merged delegated and application permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
          }

          return Promise.resolve(resolvedApis);
        }
        catch (e) {
          return Promise.reject(e);
        }
      });
  }

  private getRequiredResourceAccessForApis(apis: string | undefined, scopeType: string, logger: Logger): RequiredResourceAccess[] {
    if (!apis) {
      return [];
    }

    const resolvedApis: RequiredResourceAccess[] = [];
    const requestedApis: string[] = apis!.split(',').map(a => a.trim());
    requestedApis.forEach(api => {
      const pos: number = api.lastIndexOf('/');
      const permissionName: string = api.substr(pos + 1);
      const servicePrincipalName: string = api.substr(0, pos);
      if (this.debug) {
        logger.logToStderr(`Resolving ${api}...`);
        logger.logToStderr(`Permission name: ${permissionName}`);
        logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
      }
      const servicePrincipal = this.items.find(sp => sp.servicePrincipalNames.indexOf(servicePrincipalName) > -1);
      if (!servicePrincipal) {
        throw `Service principal ${servicePrincipalName} not found`;
      }

      const scopesOfType = scopeType === 'Scope' ? servicePrincipal.oauth2PermissionScopes : servicePrincipal.appRoles;
      const permission = scopesOfType.find(scope => scope.value === permissionName);
      if (!permission) {
        throw `Permission ${permissionName} for service principal ${servicePrincipalName} not found`;
      }

      let resolvedApi = resolvedApis.find(a => a.resourceAppId === servicePrincipal.appId);
      if (!resolvedApi) {
        resolvedApi = {
          resourceAppId: servicePrincipal.appId,
          resourceAccess: []
        };
        resolvedApis.push(resolvedApi);
      }

      resolvedApi.resourceAccess.push({
        id: permission.id,
        type: scopeType
      });
    });

    return resolvedApis;
  }

  private configureSecret(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.withSecret) {
      return Promise.resolve(appInfo);
    }

    if (this.verbose) {
      logger.logToStderr(`Configure Azure AD app secret...`);
    }

    const secretExpirationDate = new Date();
    secretExpirationDate.setFullYear(secretExpirationDate.getFullYear() + 1);

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}/addPassword`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        passwordCredential: {
          displayName: 'Default',
          endDateTime: secretExpirationDate.toISOString()
        }
      }
    };

    return request
      .post<{ secretText: string }>(requestOptions)
      .then((password: { secretText: string; }): Promise<AppInfo> => {
        appInfo.secret = password.secretText;
        return Promise.resolve(appInfo);
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      },
      {
        option: '--multitenant'
      },
      {
        option: '-r, --redirectUris [redirectUris]'
      },
      {
        option: '-p, --platform [platform]',
        autocomplete: AadAppAddCommand.aadApplicationPlatform
      },
      {
        option: '--implicitFlow'
      },
      {
        option: '-s, --withSecret'
      },
      {
        option: '--apisDelegated [apisDelegated]'
      },
      {
        option: '--apisApplication [apisApplication]'
      },
      {
        option: '-u, --uri [uri]'
      },
      {
        option: '--scopeName [scopeName]'
      },
      {
        option: '--scopeConsentBy [scopeConsentBy]',
        autocomplete: AadAppAddCommand.aadAppScopeConsentBy
      },
      {
        option: '--scopeAdminConsentDisplayName [scopeAdminConsentDisplayName]'
      },
      {
        option: '--scopeAdminConsentDescription [scopeAdminConsentDescription]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.platform &&
      AadAppAddCommand.aadApplicationPlatform.indexOf(args.options.platform) < 0) {
      return `${args.options.platform} is not a valid value for platform. Allowed values are ${AadAppAddCommand.aadApplicationPlatform.join(', ')}`;
    }

    if (args.options.redirectUris && !args.options.platform) {
      return `When you specify redirectUris you also need to specify platform`;
    }

    if (args.options.scopeName) {
      if (!args.options.uri) {
        return `When you specify scopeName you also need to specify uri`;
      }

      if (!args.options.scopeAdminConsentDescription) {
        return `When you specify scopeName you also need to specify scopeAdminConsentDescription`;
      }

      if (!args.options.scopeAdminConsentDisplayName) {
        return `When you specify scopeName you also need to specify scopeAdminConsentDisplayName`;
      }
    }

    if (args.options.scopeConsentBy &&
      AadAppAddCommand.aadAppScopeConsentBy.indexOf(args.options.scopeConsentBy) < 0) {
      return `${args.options.scopeConsentBy} is not a valid value for scopeConsentBy. Allowed values are ${AadAppAddCommand.aadAppScopeConsentBy.join(', ')}`;
    }

    return true;
  }
}

module.exports = new AadAppAddCommand();