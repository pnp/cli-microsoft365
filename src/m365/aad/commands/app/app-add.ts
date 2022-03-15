import * as fs from 'fs';
import { v4 } from 'uuid';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { accessToken, odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import { M365RcJson } from '../../../base/M365RcJson';
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
  tenantId: string;
  secret?: string;
  // used when multiple secrets have been defined in the manifest
  // in v6 we'll remove secret from AppInfo and just use secrets
  secrets?: {
    displayName: string;
    value: string;
  }[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  apisApplication?: string;
  apisDelegated?: string;
  implicitFlow: boolean;
  manifest?: string;
  multitenant: boolean;
  name?: string;
  platform?: string;
  redirectUris?: string;
  save?: boolean;
  scopeAdminConsentDescription?: string;
  scopeAdminConsentDisplayName?: string;
  scopeConsentBy?: string;
  scopeName?: string;
  uri?: string;
  withSecret: boolean;
}

class AadAppAddCommand extends GraphCommand {
  private static aadApplicationPlatform: string[] = ['spa', 'web', 'publicClient'];
  private static aadAppScopeConsentBy: string[] = ['admins', 'adminsAndUsers'];
  private manifest: any;
  private appName: string = '';

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
      .then(appInfo => {
        // based on the assumption that we're adding AAD app to the current
        // directory. If we in the future extend the command with allowing
        // users to create AAD app in a different directory, we'll need to
        // adjust this
        appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken);
        return Promise.resolve(appInfo);
      })
      .then(appInfo => this.updateAppFromManifest(args, appInfo))
      .then(appInfo => this.configureUri(args, appInfo, logger))
      .then(appInfo => this.configureSecret(args, appInfo, logger))
      .then(appInfo => this.saveAppInfo(args, appInfo, logger))
      .then((_appInfo: AppInfo): void => {
        const appInfo: any = {
          appId: _appInfo.appId,
          objectId: _appInfo.id,
          tenantId: _appInfo.tenantId
        };
        if (_appInfo.secret) {
          appInfo.secret = _appInfo.secret;
        }
        if (_appInfo.secrets) {
          appInfo.secrets = _appInfo.secrets;
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

    if (!applicationInfo.displayName && this.manifest) {
      applicationInfo.displayName = this.manifest.name;
    }
    this.appName = applicationInfo.displayName;

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

  private updateAppFromManifest(args: CommandArgs, appInfo: AppInfo): Promise<AppInfo> {
    if (!args.options.manifest) {
      return Promise.resolve(appInfo);
    }

    const v2Manifest: any = JSON.parse(args.options.manifest);
    // remove properties that might be coming from the original app that was
    // used to create the manifest and which can't be updated
    delete v2Manifest.id;
    delete v2Manifest.appId;
    delete v2Manifest.publisherDomain;
    // extract secrets from the manifest. Store them in a separate variable
    // and remove them from the manifest because we need to create them
    // separately
    const secrets: { name: string, expirationDate: Date }[] = this.getSecretsFromManifest(v2Manifest);
    // Azure Portal returns v2 manifest whereas the Graph API expects a v1.6
    const graphManifest = this.transformManifest(v2Manifest);

    const updateAppRequestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: graphManifest
    };

    return request
      .patch(updateAppRequestOptions)
      .then(_ => this.updatePreAuthorizedAppsFromManifest(v2Manifest, appInfo))
      .then(_ => this.createSecrets(secrets, appInfo));
  }

  private getSecretsFromManifest(manifest: any): { name: string, expirationDate: Date }[] {
    if (!manifest.passwordCredentials || manifest.passwordCredentials.length === 0) {
      return [];
    }

    const secrets = manifest.passwordCredentials.map((c: any) => {
      const startDate = new Date(c.startDate);
      const endDate = new Date(c.endDate);
      const expirationDate = new Date();
      expirationDate.setMilliseconds(endDate.valueOf() - startDate.valueOf());

      return {
        name: c.displayName,
        expirationDate
      };
    });

    // delete the secrets from the manifest so that we won't try to set them
    // from the manifest
    delete manifest.passwordCredentials;

    return secrets;
  }

  private updatePreAuthorizedAppsFromManifest(manifest: any, appInfo: AppInfo): Promise<AppInfo> {
    if (!manifest ||
      !manifest.preAuthorizedApplications ||
      manifest.preAuthorizedApplications.length === 0) {
      return Promise.resolve(appInfo);
    }

    const graphManifest: any = {
      api: {
        preAuthorizedApplications: manifest.preAuthorizedApplications
      }
    };

    graphManifest.api.preAuthorizedApplications.forEach((p: any) => {
      p.delegatedPermissionIds = p.permissionIds;
      delete p.permissionIds;
    });

    const updateAppRequestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: graphManifest
    };

    return request
      .patch(updateAppRequestOptions)
      .then(_ => Promise.resolve(appInfo));
  }

  private createSecrets(secrets: { name: string, expirationDate: Date }[], appInfo: AppInfo): Promise<AppInfo> {
    if (secrets.length === 0) {
      return Promise.resolve(appInfo);
    }

    return Promise
      .all(secrets.map(secret => this.createSecret({
        appObjectId: appInfo.id,
        displayName: secret.name,
        expirationDate: secret.expirationDate
      })))
      .then(secrets => {
        appInfo.secrets = secrets;
        return appInfo;
      });
  }

  private transformManifest(v2Manifest: any): any {
    const graphManifest = JSON.parse(JSON.stringify(v2Manifest));
    // add missing properties
    if (!graphManifest.api) {
      graphManifest.api = {};
    }
    if (!graphManifest.info) {
      graphManifest.info = {};
    }
    if (!graphManifest.web) {
      graphManifest.web = {
        implicitGrantSettings: {},
        redirectUris: []
      };
    }
    if (!graphManifest.spa) {
      graphManifest.spa = {
        redirectUris: []
      };
    }

    // remove properties that have no equivalent in v1.6
    const unsupportedProperties = [
      'accessTokenAcceptedVersion',
      'disabledByMicrosoftStatus',
      'errorUrl',
      'oauth2RequirePostResponse',
      'oauth2AllowUrlPathMatching',
      'orgRestrictions',
      'samlMetadataUrl'
    ];
    unsupportedProperties.forEach(p => delete graphManifest[p]);

    graphManifest.api.acceptMappedClaims = v2Manifest.acceptMappedClaims;
    delete graphManifest.acceptMappedClaims;

    graphManifest.isFallbackPublicClient = v2Manifest.allowPublicClient;
    delete graphManifest.allowPublicClient;

    graphManifest.info.termsOfServiceUrl = v2Manifest.informationalUrls?.termsOfService;
    graphManifest.info.supportUrl = v2Manifest.informationalUrls?.support;
    graphManifest.info.privacyStatementUrl = v2Manifest.informationalUrls?.privacy;
    graphManifest.info.marketingUrl = v2Manifest.informationalUrls?.marketing;
    delete graphManifest.informationalUrls;

    graphManifest.api.knownClientApplications = v2Manifest.knownClientApplications;
    delete graphManifest.knownClientApplications;

    graphManifest.info.logoUrl = v2Manifest.logoUrl;
    delete graphManifest.logoUrl;

    graphManifest.web.logoutUrl = v2Manifest.logoutUrl;
    delete graphManifest.logoutUrl;

    graphManifest.displayName = v2Manifest.name;
    delete graphManifest.name;

    graphManifest.web.implicitGrantSettings.enableAccessTokenIssuance = v2Manifest.oauth2AllowImplicitFlow;
    delete graphManifest.oauth2AllowImplicitFlow;

    graphManifest.web.implicitGrantSettings.enableIdTokenIssuance = v2Manifest.oauth2AllowIdTokenImplicitFlow;
    delete graphManifest.oauth2AllowIdTokenImplicitFlow;

    graphManifest.api.oauth2PermissionScopes = v2Manifest.oauth2Permissions;
    delete graphManifest.oauth2Permissions;
    if (graphManifest.api.oauth2PermissionScopes) {
      graphManifest.api.oauth2PermissionScopes.forEach((scope: any) => {
        delete scope.lang;
        delete scope.origin;
      });
    }

    delete graphManifest.oauth2RequiredPostResponse;

    // MS Graph doesn't support creating OAuth2 permissions and pre-authorized
    // apps in one request. This is why we need to remove it here and do it in
    // the next request
    delete graphManifest.preAuthorizedApplications;

    if (v2Manifest.replyUrlsWithType) {
      v2Manifest.replyUrlsWithType.forEach((urlWithType: any) => {
        if (urlWithType.type === 'Web') {
          graphManifest.web.redirectUris.push(urlWithType.url);
          return;
        }
        if (urlWithType.type === 'Spa') {
          graphManifest.spa.redirectUris.push(urlWithType.url);
          return;
        }
      });
      delete graphManifest.replyUrlsWithType;
    }

    graphManifest.web.homePageUrl = v2Manifest.signInUrl;
    delete graphManifest.signInUrl;

    if (graphManifest.appRoles) {
      graphManifest.appRoles.forEach((role: any) => {
        delete role.lang;
      });
    }

    return graphManifest;
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

    return odata
      .getAllItems<ServicePrincipalInfo>(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=servicePrincipalNames,appId,oauth2PermissionScopes,appRoles`, logger)
      .then(servicePrincipals => {
        try {
          const resolvedApis = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.apisDelegated, 'Scope', logger);
          if (this.debug) {
            logger.logToStderr(`Resolved delegated permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
          }
          const resolvedApplicationApis = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.apisApplication, 'Role', logger);
          if (this.debug) {
            logger.logToStderr(`Resolved application permissions: ${JSON.stringify(resolvedApplicationApis, null, 2)}`);
          }
          // merge resolved application APIs onto resolved delegated APIs
          resolvedApplicationApis.forEach(resolvedRequiredResource => {
            const requiredResource = resolvedApis.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
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

  private getRequiredResourceAccessForApis(servicePrincipals: ServicePrincipalInfo[], apis: string | undefined, scopeType: string, logger: Logger): RequiredResourceAccess[] {
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
      const servicePrincipal = servicePrincipals.find(sp => (
        sp.servicePrincipalNames.indexOf(servicePrincipalName) > -1 ||
        sp.servicePrincipalNames.indexOf(`${servicePrincipalName}/`) > -1));
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

    return this
      .createSecret({ appObjectId: appInfo.id })
      .then(secret => {
        appInfo.secret = secret.value;
        return Promise.resolve(appInfo);
      });
  }

  private createSecret({ appObjectId, displayName = undefined, expirationDate = undefined }: { appObjectId: string, displayName?: string, expirationDate?: Date }): Promise<{ displayName: string, value: string }> {
    let secretExpirationDate = expirationDate;
    if (!secretExpirationDate) {
      secretExpirationDate = new Date();
      secretExpirationDate.setFullYear(secretExpirationDate.getFullYear() + 1);
    }

    const secretName = displayName ?? 'Default';

    const requestOptions: any = {
      url: `${this.resource}/v1.0/myorganization/applications/${appObjectId}/addPassword`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        passwordCredential: {
          displayName: secretName,
          endDateTime: secretExpirationDate.toISOString()
        }
      }
    };

    return request
      .post<{ secretText: string }>(requestOptions)
      .then((password: { secretText: string; }) => Promise.resolve({
        displayName: secretName,
        value: password.secretText
      }));
  }

  private saveAppInfo(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.save) {
      return Promise.resolve(appInfo);
    }

    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      logger.logToStderr(`Saving Azure AD app registration information to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      if (this.debug) {
        logger.logToStderr(`Reading existing ${filePath}...`);
      }

      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
        return Promise.resolve(appInfo);
      }
    }

    if (!m365rc.apps) {
      m365rc.apps = [];
    }

    m365rc.apps.push({
      appId: appInfo.appId,
      name: this.appName
    });

    try {
      fs.writeFileSync(filePath, JSON.stringify(m365rc, null, 2));
    }
    catch (e) {
      logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
    }

    return Promise.resolve(appInfo);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name [name]'
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
      },
      {
        option: '--manifest [manifest]'
      },
      {
        option: '--save'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.manifest && !args.options.name) {
      return 'Specify either the name of the app to create or the manifest';
    }

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

    if (args.options.manifest) {
      try {
        this.manifest = JSON.parse(args.options.manifest);
        if (!args.options.name && !this.manifest.name) {
          return `Specify the name of the app to create either through the 'name' option or the 'name' property in the manifest`;
        }
      }
      catch (e) {
        return `Error while parsing the specified manifest: ${e}`;
      }
    }

    return true;
  }
}

module.exports = new AadAppAddCommand();