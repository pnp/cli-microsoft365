import fs from 'fs';
import forge from 'node-forge';
import auth from '../../../../Auth.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { AppInfo, AppPermissions, RequiredResourceAccess, ResourceAccess, ServicePrincipalInfo, entraApp } from '../../../../utils/entraApp.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  mode: string;
  permissions?: string;
  permissionSet?: string;
}

const PermissionScopes = {
  Delegated: {
    Full:
      [
        "https://management.azure.com/user_impersonation",
        "https://graph.microsoft.com/AppCatalog.ReadWrite.All",
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
        "https://graph.microsoft.com/IdentityProvider.ReadWrite.All",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/Reports.Read.All",
        "https://graph.microsoft.com/Tasks.ReadWrite",
        "https://graph.microsoft.com/TeamsAppInstallation.ReadWriteForUser",
        "https://graph.microsoft.com/User.Invite.All",
        "https://manage.office.com/ServiceHealth.Read",
        "https://microsoft.sharepoint-df.com/AllSites.FullControl",
        "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All",
        "https://microsoft.sharepoint-df.com/User.Read.All",
        "https://api.yammer.com/user_impersonation"
      ],
    ReadAll: [
      "https://management.azure.com/user_impersonation",
      "https://graph.microsoft.com/AppCatalog.Read.All",
      "https://graph.microsoft.com/Directory.AccessAsUser.All",
      "https://graph.microsoft.com/Directory.Read.All",
      "https://graph.microsoft.com/Group.Read.All",
      "https://graph.microsoft.com/IdentityProvider.Read.All",
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/Reports.Read.All",
      "https://graph.microsoft.com/Tasks.Read",
      "https://graph.microsoft.com/TeamsAppInstallation.ReadForUser",
      "https://manage.office.com/ServiceHealth.Read",
      "https://microsoft.sharepoint-df.com/AllSites.Read",
      "https://microsoft.sharepoint-df.com/TermStore.Read.All",
      "https://microsoft.sharepoint-df.com/User.Read.All",
      "https://api.yammer.com/user_impersonation"
    ],
    SpoFull: [
      "https://microsoft.sharepoint-df.com/AllSites.FullControl",
      "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All",
      "https://microsoft.sharepoint-df.com/User.Read.All"
    ],
    SpoRead: [
      "https://microsoft.sharepoint-df.com/AllSites.Read",
      "https://microsoft.sharepoint-df.com/TermStore.Read.All",
      "https://microsoft.sharepoint-df.com/User.Read.All"
    ]
  },
  AppOnly: {
    Full:
      [
        "https://graph.microsoft.com/AppCatalog.ReadWrite.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
        "https://graph.microsoft.com/IdentityProvider.ReadWrite.All",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/Reports.Read.All",
        "https://graph.microsoft.com/Tasks.ReadWrite.All",
        "https://graph.microsoft.com/TeamsAppInstallation.ReadWriteForUser.All",
        "https://graph.microsoft.com/User.Invite.All",
        "https://manage.office.com/ServiceHealth.Read",
        "https://microsoft.sharepoint-df.com/Sites.FullControl.All",
        "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All",
        "https://microsoft.sharepoint-df.com/User.Read.All"
      ],
    ReadAll: [
      "https://graph.microsoft.com/AppCatalog.Read.All",
      "https://graph.microsoft.com/Directory.Read.All",
      "https://graph.microsoft.com/Group.Read.All",
      "https://graph.microsoft.com/IdentityProvider.Read.All",
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/Reports.Read.All",
      "https://graph.microsoft.com/Tasks.Read.All",
      "https://graph.microsoft.com/TeamsAppInstallation.ReadForUser.All",
      "https://manage.office.com/ServiceHealth.Read",
      "https://microsoft.sharepoint-df.com/Sites.Read.All",
      "https://microsoft.sharepoint-df.com/TermStore.Read.All",
      "https://microsoft.sharepoint-df.com/User.Read.All"
    ],
    SpoFull: [
      "https://microsoft.sharepoint-df.com/Sites.FullControl.All",
      "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All",
      "https://microsoft.sharepoint-df.com/User.Read.All"
    ],
    SpoRead: [
      "https://microsoft.sharepoint-df.com/Sites.Read.All",
      "https://microsoft.sharepoint-df.com/TermStore.Read.All",
      "https://microsoft.sharepoint-df.com/User.Read.All"
    ]
  }
};

class CliAppAddCommand extends GraphCommand {
  private static mode: string[] = ['delegated', 'appOnly'];
  private static permissionSet: string[] = ['ReadAll', 'SpoFull', 'SpoRead', 'Full'];
  private appPermissions: AppPermissions[] = [];

  public get name(): string {
    return commands.APP_ADD;
  }

  public get description(): string {
    return 'Creates custom Entra Id app for use by CLI for Microsoft 365';
  }

  constructor() {
    super();
    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        name: typeof args.options.name !== 'undefined',
        mode: typeof args.options.mode !== 'undefined',
        permissions: typeof args.options.permissions !== 'undefined',
        permissionSet: typeof args.options.permissionSet !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-m, --mode <mode>',
        autocomplete: CliAppAddCommand.mode
      },
      { option: '-p, --permissions [permissions]' },
      {
        option: '--permissionSet [permissionSet]',
        autocomplete: CliAppAddCommand.permissionSet
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.mode &&
          CliAppAddCommand.mode.indexOf(args.options.mode) < 0) {
          return `${args.options.mode} is not a valid value for 'mode' option. Allowed values are ${CliAppAddCommand.mode.join(', ')}`;
        }

        if (args.options.permissions && args.options.permissionSet) {
          return `Specify either 'permissions' or 'permissionSet, but not both.`;
        }

        if (!args.options.permissions && !args.options.permissionSet) {
          return `Either 'permissions' or 'permissionSet option must be specified.`;
        }

        if (args.options.permissionSet &&
          CliAppAddCommand.permissionSet.indexOf(args.options.permissionSet) < 0) {
          return `${args.options.permissionSet} is not a valid value for 'permissionSet'. Allowed values are ${CliAppAddCommand.permissionSet.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const permissions = args.options.permissions ? args.options.permissions : await this.getCorrectPermissionSet(args, logger);
      const apis = await this.resolveApis(permissions, args.options.mode, logger);
      const password: string | undefined = this.generatePassword();
      const appInfo: any = await this.createAppRegistration(args, apis, password, logger);
      await this.grantAdminConsent(appInfo, logger);

      const result: any = {
        appId: appInfo.appId,
        objectId: appInfo.id,
        tenantId: accessToken.getTenantIdFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken),
        name: appInfo.displayName
      };

      if (args.options.mode === 'appOnly') {
        result.certPassword = password;
        result.certThumbprint = appInfo.keyCredentials && appInfo.keyCredentials.length > 0 ? appInfo.keyCredentials[0].customKeyIdentifier : undefined;
        result.certExpirationDate = appInfo.keyCredentials && appInfo.keyCredentials.length > 0 ? appInfo.keyCredentials[0].endDateTime : undefined;
      }

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private generatePassword(length: number = 32): string {
    const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let password = "";
    for (let i = 0; i < length; ++i) {
      const randomIndex = Math.floor(Math.random() * charset.length);
      password += charset.charAt(randomIndex);
    }
    return password;
  }

  private async getCorrectPermissionSet(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.debug) {
      await logger.logToStderr(`Getting a correct permissionSet...`);
    }

    const permissionSetGroup = args.options.mode === 'delegated' ? PermissionScopes.Delegated : PermissionScopes.AppOnly;
    let permissionSet: string[] = permissionSetGroup.Full;

    if (args.options.permissionSet) {
      switch (args.options.permissionSet) {
        case 'ReadAll':
          permissionSet = permissionSetGroup.ReadAll;
          break;
        case 'SpoFull':
          permissionSet = permissionSetGroup.SpoFull;
          break;
        case 'SpoRead':
          permissionSet = permissionSetGroup.SpoRead;
          break;
        default:
          permissionSet = permissionSetGroup.Full;
      }
    }

    return permissionSet.join(',');
  }

  private async createAppRegistration(args: CommandArgs, apis: RequiredResourceAccess[], password: string, logger: Logger): Promise<AppInfo> {
    const applicationInfo: any = {
      displayName: args.options.name,
      signInAudience: 'AzureADMyOrg',
      publicClient: {
        redirectUris: ['https://login.microsoftonline.com/common/oauth2/nativeclient']
      },
      isFallbackPublicClient: true
    };

    if (apis.length > 0) {
      applicationInfo.requiredResourceAccess = apis;
    }

    if (args.options.mode === 'appOnly') {
      const certificateBase64Encoded = await this.generateCertificateBase64Encoded(password, logger);

      applicationInfo.keyCredentials = [{
        type: "AsymmetricX509Cert",
        usage: "Verify",
        displayName: 'PnP M365 Management Shell',
        key: certificateBase64Encoded
      }];
    }

    if (this.verbose) {
      await logger.logToStderr(`Creating Microsoft Entra app registration...`);
    }

    return await entraApp.createEntraApp(applicationInfo);
  }

  private async generateCertificateBase64Encoded(password: string, logger: Logger): Promise<string> {
    if (this.debug) {
      await logger.logToStderr(`Creating a new certificate...`);
    }

    try {
      const { pki, pkcs12, asn1, util } = forge;
      const cert = pki.createCertificate();
      const keys = pki.rsa.generateKeyPair(2048);

      cert.validity.notBefore = new Date();
      cert.validity.notAfter = new Date();
      cert.validity.notAfter.setFullYear(cert.validity.notBefore.getFullYear() + 1);
      cert.setSubject([{ name: 'commonName', value: 'PnP M365 Management Shell' }]);
      cert.setIssuer([{ name: 'commonName', value: 'PnP M365 Management Shell' }]);
      cert.publicKey = keys.publicKey;
      cert.sign(keys.privateKey);

      const pemCert = pki.certificateToPem(cert);
      fs.writeFileSync('PnP-Certificate.cer', pemCert);

      const p12Asn1 = pkcs12.toPkcs12Asn1(keys.privateKey, cert, password);
      const p12Der = asn1.toDer(p12Asn1).getBytes();

      fs.writeFileSync('PnP-Certificate.pfx', p12Der, 'binary');

      return util.encode64(pemCert);
    }
    catch (e) {
      throw new Error(`Error while creating certificate file: ${e}.`);
    }
  }

  private async resolveApis(permissions: string, mode: string, logger: Logger): Promise<RequiredResourceAccess[]> {
    if (this.verbose) {
      await logger.logToStderr('Resolving requested APIs...');
    }

    const servicePrincipals = await odata.getAllItems<ServicePrincipalInfo>(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);

    try {
      const scopeType = mode === 'delegated' ? 'Scope' : 'Role';
      const resolvedApis: RequiredResourceAccess[] = await this.getRequiredResourceAccessForApis(servicePrincipals, permissions, scopeType, logger);
      if (this.verbose) {
        await logger.logToStderr(`Resolved ${mode} permissions: ${JSON.stringify(resolvedApis, null, 2)}`);
      }

      return resolvedApis;
    }
    catch (e) {
      throw e;
    }
  }

  private async grantAdminConsent(appInfo: AppInfo, logger: Logger): Promise<void> {
    const sp = await entraApp.createServicePrincipal(appInfo.appId);
    if (this.debug) {
      await logger.logToStderr("Service principal created, returned object id: " + sp.id);
    }

    const tasks: Promise<void>[] = [];

    this.appPermissions.forEach(async (permission) => {
      if (permission.scope.length > 0) {
        tasks.push(entraApp.grantOAuth2Permission(sp.id, permission.resourceId, permission.scope.join(' ')));

        if (this.debug) {
          await logger.logToStderr(`Admin consent granted for following resource ${permission.resourceId}, with delegated permissions: ${permission.scope.join(',')}`);
        }
      }

      permission.resourceAccess.filter(access => access.type === "Role").forEach(async (access: ResourceAccess) => {
        tasks.push(entraApp.addRoleToServicePrincipal(sp.id, permission.resourceId, access.id));

        if (this.debug) {
          await logger.logToStderr(`Admin consent granted for following resource ${permission.resourceId}, with application permission: ${access.id}`);
        }
      });
    });

    await Promise.all(tasks);
  }

  private async getRequiredResourceAccessForApis(servicePrincipals: ServicePrincipalInfo[], apis: string, scopeType: string, logger: Logger): Promise<RequiredResourceAccess[]> {
    const resolvedApis: RequiredResourceAccess[] = [];
    const requestedApis: string[] = apis!.split(',').map(a => a.trim());
    for (const api of requestedApis) {
      const pos: number = api.lastIndexOf('/');
      const permissionName: string = api.substr(pos + 1);
      const servicePrincipalName: string = api.substr(0, pos);
      if (this.debug) {
        await logger.logToStderr(`Resolving ${api}...`);
        await logger.logToStderr(`Permission name: ${permissionName}`);
        await logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
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

      const resourceAccessPermission = {
        id: permission.id,
        type: scopeType
      };

      resolvedApi.resourceAccess.push(resourceAccessPermission);

      this.updateAppPermissions(servicePrincipal.id, resourceAccessPermission, permission.value);
    }

    return resolvedApis;
  }

  private updateAppPermissions(spId: string, resourceAccessPermission: ResourceAccess, oAuth2PermissionValue?: string): void {
    // During API resolution, we store globally both app role assignments and oauth2permissions
    // So that we'll be able to parse them during the admin consent process
    let existingPermission = this.appPermissions.find(oauth => oauth.resourceId === spId);
    if (!existingPermission) {
      existingPermission = {
        resourceId: spId,
        resourceAccess: [],
        scope: []
      };

      this.appPermissions.push(existingPermission);
    }

    if (resourceAccessPermission.type === 'Scope' && oAuth2PermissionValue && !existingPermission.scope.find(scp => scp === oAuth2PermissionValue)) {
      existingPermission.scope.push(oAuth2PermissionValue);
    }

    if (!existingPermission.resourceAccess.find(res => res.id === resourceAccessPermission.id)) {
      existingPermission.resourceAccess.push(resourceAccessPermission);
    }
  }
}
export default new CliAppAddCommand();