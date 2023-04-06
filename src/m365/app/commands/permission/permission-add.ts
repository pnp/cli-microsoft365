import { Application, AppRole, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import AppCommand from '../../../base/AppCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  applicationPermission?: string;
  delegatedPermission?: string;
  grantAdminConsent?: boolean;
}

interface AppPermission {
  resourceId: string;
  resourceAccess: ResourceAccess[];
  scope: string[];
}

enum ScopeType {
  Role = 'Role',
  Scope = 'Scope'
}

class AppPermissionAddCommand extends AppCommand {
  public get name(): string {
    return commands.PERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds the specified application and/or delegated permissions to the current AAD app API permissions';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        applicationPermission: typeof args.options.applicationPermission !== 'undefined',
        delegatedPermission: typeof args.options.delegatedPermission !== 'undefined',
        grantAdminConsent: !!args.options.grantAdminConsent
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '--appId [appId]' },
      { option: '--applicationPermission [applicationPermission]' },
      { option: '--delegatedPermission [delegatedPermission]' },
      { option: '--grantAdminConsent' }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['applicationPermission', 'delegatedPermission'],
      runsWhen: (args) => args.options.delegatedPermission === undefined && args.options.applicationPermission === undefined
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appObject = await this.getAppObject();
      const servicePrincipals = await odata.getAllItems<ServicePrincipal>(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);

      let delegatedPermissions: RequiredResourceAccess[] = [], applicationPermissions: RequiredResourceAccess[] = [];
      const appPermissions: AppPermission[] = [];

      if (args.options.delegatedPermission) {
        delegatedPermissions = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermission, ScopeType.Scope, appPermissions, logger);
        this.addPermissionsToResourceArray(delegatedPermissions, appObject.requiredResourceAccess!);
      }
      if (args.options.applicationPermission) {
        applicationPermissions = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermission, ScopeType.Role, appPermissions, logger);
        this.addPermissionsToResourceArray(applicationPermissions, appObject.requiredResourceAccess!);
      }

      const addPermissionsRequestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/myorganization/applications/${appObject.id}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          requiredResourceAccess: appObject.requiredResourceAccess
        }
      };
      await request.patch(addPermissionsRequestOptions);

      if (args.options.grantAdminConsent) {
        const appServicePrincipal = servicePrincipals.find(sp => sp.appId === this.appId);
        await this.grantAdminConsent(appServicePrincipal!, appPermissions, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObject(): Promise<Application> {
    const apps = await odata.getAllItems<Application>(`${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(this.appId!)}'&$select=id,requiredResourceAccess`);
    if (apps.length === 0) {
      throw `App with id ${this.appId} not found in Azure Active Directory`;
    }
    return apps[0];
  }

  private addPermissionsToResourceArray(permissions: RequiredResourceAccess[], existingArray: RequiredResourceAccess[]): void {
    permissions.forEach(resolvedRequiredResource => {
      const requiredResource = existingArray.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
      if (requiredResource) {
        // make sure that permission does not yet exist on the app or it will be added twice
        resolvedRequiredResource.resourceAccess!.forEach(resAccess => {
          if (!requiredResource.resourceAccess!.some(res => res.id === resAccess.id)) {
            requiredResource.resourceAccess!.push(resAccess);
          }
        });
      }
      else {
        existingArray.push(resolvedRequiredResource);
      }
    });
  }

  private getRequiredResourceAccessForApis(servicePrincipals: ServicePrincipal[], apis: string, scopeType: string, appPermissions: AppPermission[], logger: Logger): RequiredResourceAccess[] {
    const resolvedApis: RequiredResourceAccess[] = [];
    const requestedApis: string[] = apis.split(' ').map(a => a.trim());

    requestedApis.forEach(api => {
      const pos: number = api.lastIndexOf('/');
      const permissionName: string = api.substring(pos + 1);
      const servicePrincipalName: string = api.substring(0, pos);
      if (this.verbose) {
        logger.logToStderr(`Resolving ${api}...`);
        logger.logToStderr(`Permission name: ${permissionName}`);
        logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
      }
      const servicePrincipal = servicePrincipals.find(sp => (
        sp.servicePrincipalNames!.indexOf(servicePrincipalName) > -1 ||
        sp.servicePrincipalNames!.indexOf(`${servicePrincipalName}/`) > -1));

      if (!servicePrincipal) {
        throw `Service principal ${servicePrincipalName} not found`;
      }

      let permission: AppRole | PermissionScope | undefined = undefined;

      if (scopeType === ScopeType.Scope) {
        permission = servicePrincipal.oauth2PermissionScopes!.find(scope => scope.value === permissionName);
      }
      else if (scopeType === ScopeType.Role) {
        permission = servicePrincipal.appRoles!.find(scope => scope.value === permissionName);
      }

      if (!permission) {
        throw `Permission ${permissionName} for service principal ${servicePrincipalName} not found`;
      }

      let resolvedApi = resolvedApis.find(a => a.resourceAppId === servicePrincipal.appId);
      if (!resolvedApi) {
        resolvedApi = {
          resourceAppId: servicePrincipal.appId!,
          resourceAccess: []
        };
        resolvedApis.push(resolvedApi);
      }

      const resourceAccessPermission = {
        id: permission.id,
        type: scopeType
      };
      resolvedApi.resourceAccess!.push(resourceAccessPermission);

      this.updateAppPermissions(servicePrincipal.id!, resourceAccessPermission, permission.value!, appPermissions);
    });
    return resolvedApis;
  }

  private updateAppPermissions(spId: string, resourceAccessPermission: ResourceAccess, oAuth2PermissionValue: string, appPermissions: AppPermission[]): void {
    let existingPermission = appPermissions.find(oauth => oauth.resourceId === spId);
    if (!existingPermission) {
      existingPermission = {
        resourceId: spId,
        resourceAccess: [],
        scope: []
      };

      appPermissions.push(existingPermission);
    }

    if (resourceAccessPermission.type === ScopeType.Scope && oAuth2PermissionValue && !existingPermission.scope.find(scp => scp === oAuth2PermissionValue)) {
      existingPermission.scope.push(oAuth2PermissionValue);
    }

    if (!existingPermission.resourceAccess.find(res => res.id === resourceAccessPermission.id)) {
      existingPermission.resourceAccess.push(resourceAccessPermission);
    }
  }

  private async grantAdminConsent(servicePrincipal: ServicePrincipal, appPermissions: AppPermission[], logger: Logger): Promise<void> {
    for await (const permission of appPermissions) {
      if (permission.scope.length > 0) {
        if (this.verbose) {
          logger.logToStderr(`Granting consent for delegated permission(s) with resourceId ${permission.resourceId} and scope(s) ${permission.scope.join(' ')}`);
        }
        await this.grantOAuth2Permission(servicePrincipal.id!, permission.resourceId!, permission.scope.join(' '));
      }

      for await (const access of permission.resourceAccess.filter(acc => acc.type === ScopeType.Role)) {
        if (this.verbose) {
          logger.logToStderr(`Granting consent for application permission with resourceId ${permission.resourceId} and appRoleId ${access.id}`);
        }
        await this.addRoleToServicePrincipal(servicePrincipal.id!, permission.resourceId, access.id!);
      }
    }
  }

  private async grantOAuth2Permission(servicePricipalId: string, resourceId: string, scope: string): Promise<void> {
    const grantAdminConsentApplicationRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/oauth2PermissionGrants`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        clientId: servicePricipalId,
        consentType: 'AllPrincipals',
        principalId: null,
        resourceId: resourceId,
        scope: scope
      }
    };

    return request.post<void>(grantAdminConsentApplicationRequestOptions);
  }

  private async addRoleToServicePrincipal(servicePrincipalId: string, resourceId: string, appRoleId: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/servicePrincipals/${servicePrincipalId}/appRoleAssignments`,
      headers: {
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        appRoleId: appRoleId,
        principalId: servicePrincipalId,
        resourceId: resourceId
      }
    };

    return request.post<void>(requestOptions);
  }
}

module.exports = new AppPermissionAddCommand();