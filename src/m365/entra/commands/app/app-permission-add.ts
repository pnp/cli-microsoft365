import { AppRole, Application, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import { odata } from "../../../../utils/odata.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import aadCommands from "../../aadCommands.js";
import { formatting } from "../../../../utils/formatting.js";
import { cli } from "../../../../cli/cli.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appName?: string;
  appObjectId?: string;
  applicationPermissions?: string;
  delegatedPermissions?: string;
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

class EntraAppPermissionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds the specified application and/or delegated permissions to a specified Entra application registration';
  }

  public alias(): string[] | undefined {
    return [aadCommands.APP_PERMISSION_ADD, commands.APPREGISTRATION_PERMISSION_ADD];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined',
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        applicationPermissions: typeof args.options.applicationPermissions !== 'undefined',
        delegatedPermissions: typeof args.options.delegatedPermissions !== 'undefined',
        grantAdminConsent: !!args.options.grantAdminConsent
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --appId [appId]' },
      { option: '-n, --appName [appName]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '-a, --applicationPermissions [applicationPermissions]' },
      { option: '-d, --delegatedPermissions [delegatedPermissions]' },
      { option: '--grantAdminConsent' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId as string)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId as string)) {
          return `${args.options.appObjectId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appName', 'appObjectId'] });
    this.optionSets.push({
      options: ['applicationPermissions', 'delegatedPermissions'],
      runsWhen: (args) => args.options.delegatedPermissions === undefined && args.options.applicationPermissions === undefined
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appObject = await this.getAppObject(args.options);
      const servicePrincipals = await this.getServicePrincipals();
      const appPermissions: AppPermission[] = [];

      if (args.options.delegatedPermissions) {
        const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
        this.addPermissionsToResourceArray(delegatedPermissions, appObject.requiredResourceAccess!);
      }

      if (args.options.applicationPermissions) {
        const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
        this.addPermissionsToResourceArray(applicationPermissions, appObject.requiredResourceAccess!);
      }

      const addPermissionsRequestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/applications/${appObject.id}`,
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
        const appServicePrincipal = servicePrincipals.find(sp => sp.appId === appObject.appId);
        await this.grantAdminConsent(appServicePrincipal!, appPermissions, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectByName(appName: string): Promise<Application[]> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=displayName eq '${formatting.encodeQueryParameter(appName)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 1) {
      return await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=id eq '${res.value[0].id}'&$select=id,appId,requiredResourceAccess`);
    }

    if (res.value.length === 0) {
      throw `No Entra application registration with name ${appName} found`;
    }

    const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', res.value);
    const result = await cli.handleMultipleResultsFound<{ id: string }>(`Multiple Entra application registration with name '${appName}' found.`, resultAsKeyValuePair);

    return await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=id eq '${result.id}'&$select=id,appId,requiredResourceAccess`);
  }

  private async getAppObject(options: Options): Promise<Application> {
    let appNotFoundMessage = '';
    let apps: Application[] = [];

    if (options.appId) {
      apps = await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=appId eq '${options.appId}'&$select=id,appId,requiredResourceAccess`);
      appNotFoundMessage = `client id ${options.appId}`;
    }
    else if (options.appName) {
      apps = await this.getAppObjectByName(options.appName);
      appNotFoundMessage = `name ${options.appName}`;
    }
    else if (options.appObjectId) {
      apps = await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=id eq '${options.appObjectId}'&$select=id,appId,requiredResourceAccess`);
      appNotFoundMessage = `object id ${options.appObjectId}`;
    }

    if (apps.length === 0) {
      throw `App with ${appNotFoundMessage} not found in Entra ID (Azure AD)`;
    }

    return apps[0];
  }

  private async getServicePrincipals(): Promise<ServicePrincipal[]> {
    return await odata.getAllItems<ServicePrincipal>(`${this.resource}/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);
  }

  private async grantAdminConsent(servicePrincipal: ServicePrincipal, appPermissions: AppPermission[], logger: Logger): Promise<void> {
    for await (const permission of appPermissions) {
      if (permission.scope.length > 0) {
        if (this.verbose) {
          await logger.logToStderr(`Granting consent for delegated permission(s) with resourceId ${permission.resourceId} and scope(s) ${permission.scope.join(' ')}`);
        }

        await this.grantOAuth2Permission(servicePrincipal.id!, permission.resourceId!, permission.scope.join(' '));
      }

      for await (const access of permission.resourceAccess.filter(acc => acc.type === ScopeType.Role)) {
        if (this.verbose) {
          await logger.logToStderr(`Granting consent for application permission with resourceId ${permission.resourceId} and appRoleId ${access.id}`);
        }

        await this.addRoleToServicePrincipal(servicePrincipal.id!, permission.resourceId, access.id!);
      }
    }
  }

  private async grantOAuth2Permission(servicePrincipalId: string, resourceId: string, scope: string): Promise<void> {
    const grantAdminConsentApplicationRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/oauth2PermissionGrants`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        clientId: servicePrincipalId,
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
      url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`,
      headers: {
        'content-type': 'application/json;odata.metadata=none'
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

  private async getRequiredResourceAccessForApis(servicePrincipals: ServicePrincipal[], apis: string, scopeType: string, appPermissions: AppPermission[], logger: Logger): Promise<RequiredResourceAccess[]> {
    const resolvedApis: RequiredResourceAccess[] = [];
    const requestedApis: string[] = apis.split(' ').map(a => a.trim());

    for await (const api of requestedApis) {
      const pos: number = api.lastIndexOf('/');
      const permissionName: string = api.substring(pos + 1);
      const servicePrincipalName: string = api.substring(0, pos);

      if (this.verbose) {
        await logger.logToStderr(`Resolving ${api}...`);
        await logger.logToStderr(`Permission name: ${permissionName}`);
        await logger.logToStderr(`Service principal name: ${servicePrincipalName}`);
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
    }
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
}

export default new EntraAppPermissionAddCommand();