import fs from 'fs';
import { v4 } from 'uuid';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { AppCreationOptions, AppInfo, entraApp } from '../../../../utils/entraApp.js';
import { optionsUtils } from '../../../../utils/optionsUtils.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { M365RcJson } from '../../../base/M365RcJson.js';
import commands from '../../commands.js';

const entraApplicationPlatform = ['spa', 'web', 'publicClient', 'apple', 'android'] as const;
const entraAppScopeConsentBy = ['admins', 'adminsAndUsers'] as const;

const options = globalOptionsZod
  .extend({
    name: zod.alias('n', z.string().optional()),
    multitenant: z.boolean().optional(),
    redirectUris: zod.alias('r', z.string().optional()),
    platform: zod.alias('p', z.enum(entraApplicationPlatform).optional()),
    implicitFlow: z.boolean().optional(),
    withSecret: zod.alias('s', z.boolean().optional()),
    apisDelegated: z.string().optional(),
    apisApplication: z.string().optional(),
    uri: zod.alias('u', z.string().optional()),
    scopeName: z.string().optional(),
    scopeConsentBy: z.enum(entraAppScopeConsentBy).optional(),
    scopeAdminConsentDisplayName: z.string().optional(),
    scopeAdminConsentDescription: z.string().optional(),
    certificateFile: z.string().optional(),
    certificateBase64Encoded: z.string().optional(),
    certificateDisplayName: z.string().optional(),
    manifest: z.string().optional(),
    save: z.boolean().optional(),
    grantAdminConsent: z.boolean().optional(),
    allowPublicClientFlows: z.boolean().optional(),
    bundleId: z.string().optional(),
    signatureHash: z.string().optional()
  })
  .passthrough();

declare type Options = z.infer<typeof options> & AppCreationOptions;

interface CommandArgs {
  options: Options;
}

class EntraAppAddCommand extends GraphCommand {
  private manifest: any;
  private appName: string = '';

  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Creates new Entra app registration';
  }

  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => {
        if (options.redirectUris && !options.platform) {
          return false;
        }
        return true;
      }, {
        message: 'When you specify redirectUris you also need to specify platform'
      })
      .refine(options => {
        if (options.platform && ['spa', 'web', 'publicClient'].includes(options.platform) && !options.redirectUris) {
          return false;
        }
        return true;
      }, {
        message: 'When you use platform spa, web or publicClient, you\'ll need to specify redirectUris'
      })
      .refine(options => {
        if (options.certificateFile && options.certificateBase64Encoded) {
          return false;
        }
        return true;
      }, {
        message: 'Specify either certificateFile or certificateBase64Encoded but not both'
      })
      .refine(options => {
        if (options.certificateDisplayName && !options.certificateFile && !options.certificateBase64Encoded) {
          return false;
        }
        return true;
      }, {
        message: 'When you specify certificateDisplayName you also need to specify certificateFile or certificateBase64Encoded'
      })
      .refine(options => {
        if (options.certificateFile && !fs.existsSync(options.certificateFile)) {
          return false;
        }
        return true;
      }, {
        message: 'Certificate file not found'
      })
      .refine(options => {
        if (options.scopeName) {
          if (!options.uri) {
            return false;
          }
          if (!options.scopeAdminConsentDescription) {
            return false;
          }
          if (!options.scopeAdminConsentDisplayName) {
            return false;
          }
        }
        return true;
      }, {
        message: 'When you specify scopeName you also need to specify uri, scopeAdminConsentDescription, and scopeAdminConsentDisplayName'
      })
      .refine(options => {
        if (options.manifest) {
          try {
            const manifest = JSON.parse(options.manifest);
            if (!options.name && !manifest.name) {
              return false;
            }
            this.manifest = manifest;
            return true;
          }
          catch (e) {
            return false;
          }
        }
        return true;
      }, {
        message: 'Specify the name of the app to create either through the \'name\' option or the \'name\' property in the manifest'
      })
      .refine(options => options.name || options.manifest, {
        message: 'Specify either name or manifest'
      })
      .refine(options => {
        if (options.platform === 'apple' && !options.bundleId) {
          return false;
        }
        return true;
      }, {
        message: 'When you use platform apple, you\'ll need to specify bundleId'
      })
      .refine(options => {
        if (options.platform === 'android' && (!options.bundleId || !options.signatureHash)) {
          return false;
        }
        return true;
      }, {
        message: 'When you use platform android, you\'ll need to specify bundleId and signatureHash'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (!args.options.name && this.manifest) {
      args.options.name = this.manifest.name;
    }
    this.appName = args.options.name!;

    try {
      const apis = await entraApp.resolveApis({
        options: args.options,
        manifest: this.manifest,
        logger,
        verbose: this.verbose,
        debug: this.debug
      });
      let appInfo: any = await entraApp.createAppRegistration({
        options: args.options,
        unknownOptions: optionsUtils.getUnknownOptions(args.options, zod.schemaToOptions(this.schema!)),
        apis,
        logger,
        verbose: this.verbose,
        debug: this.debug
      });
      // based on the assumption that we're adding Microsoft Entra app to the current
      // directory. If we in the future extend the command with allowing
      // users to create Microsoft Entra app in a different directory, we'll need to
      // adjust this
      appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      appInfo = await this.updateAppFromManifest(args, appInfo);
      appInfo = await entraApp.grantAdminConsent({
        appInfo,
        appPermissions: entraApp.appPermissions,
        adminConsent: args.options.grantAdminConsent,
        logger,
        debug: this.debug
      });
      appInfo = await this.configureUri(args, appInfo, logger);
      appInfo = await this.configureSecret(args, appInfo, logger);
      const _appInfo = await this.saveAppInfo(args, appInfo, logger);

      appInfo = {
        appId: _appInfo.appId,
        objectId: _appInfo.id,
        tenantId: _appInfo.tenantId
      };
      if (_appInfo.secrets) {
        appInfo.secrets = _appInfo.secrets;
      }

      await logger.log(appInfo);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async configureSecret(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.withSecret || (appInfo.secrets && appInfo.secrets.length > 0)) {
      return appInfo;
    }

    if (this.verbose) {
      await logger.logToStderr(`Configure Microsoft Entra app secret...`);
    }

    const secret = await this.createSecret({ appObjectId: appInfo.id });

    if (!appInfo.secrets) {
      appInfo.secrets = [];
    }
    appInfo.secrets.push(secret);
    return appInfo;
  }

  private async createSecret({ appObjectId, displayName = undefined, expirationDate = undefined }: { appObjectId: string, displayName?: string, expirationDate?: Date }): Promise<{ displayName: string, value: string }> {
    let secretExpirationDate = expirationDate;
    if (!secretExpirationDate) {
      secretExpirationDate = new Date();
      secretExpirationDate.setFullYear(secretExpirationDate.getFullYear() + 1);
    }

    const secretName = displayName ?? 'Default';

    const requestOptions: CliRequestOptions = {
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

    const response = await request.post<{ secretText: string }>(requestOptions);
    return {
      displayName: secretName,
      value: response.secretText
    };
  }

  private async updateAppFromManifest(args: CommandArgs, appInfo: AppInfo): Promise<AppInfo> {
    if (!args.options.manifest) {
      return appInfo;
    }

    const v2Manifest: any = JSON.parse(args.options.manifest);
    // remove properties that might be coming from the original app that was
    // used to create the manifest and which can't be updated
    delete v2Manifest.id;
    delete v2Manifest.appId;
    delete v2Manifest.publisherDomain;

    // extract secrets from the manifest. Store them in a separate variable
    const secrets: { name: string, expirationDate: Date }[] = this.getSecretsFromManifest(v2Manifest);

    // Azure Portal returns v2 manifest whereas the Graph API expects a v1.6

    if (args.options.apisApplication || args.options.apisDelegated) {
      // take submitted delegated / application permissions as options
      // otherwise, they will be skipped in the app update
      v2Manifest.requiredResourceAccess = appInfo.requiredResourceAccess;
    }

    if (args.options.redirectUris) {
      // take submitted redirectUris/platform as options
      // otherwise, they will be removed from the app
      v2Manifest.replyUrlsWithType = args.options.redirectUris.split(',').map(u => {
        return {
          url: u.trim(),
          type: this.translatePlatformToType(args.options.platform!)
        };
      });
    }

    if (args.options.multitenant) {
      // override manifest setting when using multitenant flag
      v2Manifest.signInAudience = 'AzureADMultipleOrgs';
    }

    if (args.options.implicitFlow) {
      // remove manifest settings when using implicitFlow flag
      delete v2Manifest.oauth2AllowIdTokenImplicitFlow;
      delete v2Manifest.oauth2AllowImplicitFlow;
    }

    if (args.options.scopeName) {
      // override manifest setting when using options.
      delete v2Manifest.oauth2Permissions;
    }

    if (args.options.certificateFile || args.options.certificateBase64Encoded) {
      // override manifest setting when using options.
      delete v2Manifest.keyCredentials;
    }

    const graphManifest = this.transformManifest(v2Manifest);

    const updateAppRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: graphManifest
    };
    await request.patch(updateAppRequestOptions);
    await this.updatePreAuthorizedAppsFromManifest(v2Manifest, appInfo);
    await this.createSecrets(secrets, appInfo);
    return appInfo;
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

  private async updatePreAuthorizedAppsFromManifest(manifest: any, appInfo: AppInfo): Promise<AppInfo> {
    if (!manifest ||
      !manifest.preAuthorizedApplications ||
      manifest.preAuthorizedApplications.length === 0) {
      return appInfo;
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

    await request.patch(updateAppRequestOptions);
    return appInfo;
  }

  private async createSecrets(secrets: { name: string, expirationDate: Date }[], appInfo: AppInfo): Promise<AppInfo> {
    if (secrets.length === 0) {
      return appInfo;
    }

    const secretsOutput: any = await Promise
      .all(secrets.map(secret => this.createSecret({
        appObjectId: appInfo.id,
        displayName: secret.name,
        expirationDate: secret.expirationDate
      })));
    appInfo.secrets = secretsOutput;
    return appInfo;
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

  private async configureUri(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.uri) {
      return appInfo;
    }

    if (this.verbose) {
      await logger.logToStderr(`Configuring Microsoft Entra application ID URI...`);
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

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appInfo.id}`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: applicationInfo
    };

    await request.patch(requestOptions);
    return appInfo;
  }

  private async saveAppInfo(args: CommandArgs, appInfo: AppInfo, logger: Logger): Promise<AppInfo> {
    if (!args.options.save) {
      return appInfo;
    }

    const filePath: string = '.m365rc.json';

    if (this.verbose) {
      await logger.logToStderr(`Saving Microsoft Entra app registration information to the ${filePath} file...`);
    }

    let m365rc: M365RcJson = {};
    if (fs.existsSync(filePath)) {
      if (this.debug) {
        await logger.logToStderr(`Reading existing ${filePath}...`);
      }

      try {
        const fileContents: string = fs.readFileSync(filePath, 'utf8');
        if (fileContents) {
          m365rc = JSON.parse(fileContents);
        }
      }
      catch (e) {
        await logger.logToStderr(`Error reading ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
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
      await logger.logToStderr(`Error writing ${filePath}: ${e}. Please add app info to ${filePath} manually.`);
    }

    return Promise.resolve(appInfo);
  }

  private translatePlatformToType(platform: string): string {
    if (platform === 'publicClient') {
      return 'InstalledClient';
    }

    return platform.charAt(0).toUpperCase() + platform.substring(1);
  }
}

export default new EntraAppAddCommand();