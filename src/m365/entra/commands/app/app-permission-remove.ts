import { AppRole, AppRoleAssignment, Application, OAuth2PermissionGrant, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import { odata } from "../../../../utils/odata.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import { cli } from "../../../../cli/cli.js";
import { formatting } from "../../../../utils/formatting.js";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appObjectId?: string;
  appName?: string;
  applicationPermissions?: string;
  delegatedPermissions?: string;
  revokeAdminConsent?: boolean;
  force?: boolean;
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

class EntraAppPermissionRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PERMISSION_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified application and/or delegated permissions from a specified Microsoft Entra app';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        appId: typeof args.options.appId !== 'undefined',
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        appName: typeof args.options.appName !== 'undefined',
        applicationPermissions: typeof args.options.applicationPermissions !== 'undefined',
        delegatedPermissions: typeof args.options.delegatedPermissions !== 'undefined',
        revokeAdminConsent: !!args.options.revokeAdminConsent,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '--appObjectId [appObjectId]'
      },
      {
        option: '-n, --appName [appName]'
      },
      {
        option: '-a, --applicationPermissions [applicationPermissions]'
      },
      {
        option: '-d, --delegatedPermissions [delegatedPermissions]'
      },
      {
        option: '--revokeAdminConsent'
      },
      {
        option: '--force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (args.options.appObjectId && !validation.isValidGuid(args.options.appObjectId)) {
          return `${args.options.appObjectId} is not a valid GUID`;
        }

        if (args.options.delegatedPermissions) {
          const invalidPermissions = validation.isValidPermission(args.options.delegatedPermissions);
          if (Array.isArray(invalidPermissions)) {
            return `Delegated permission(s) ${invalidPermissions.join(', ')} are not fully-qualified`;
          }
        }

        if (args.options.applicationPermissions) {
          const invalidPermissions = validation.isValidPermission(args.options.applicationPermissions);
          if (Array.isArray(invalidPermissions)) {
            return `Application permission(s) ${invalidPermissions.join(', ')} are not fully-qualified`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['appId', 'appObjectId', 'appName']
      },
      {
        options: ['applicationPermissions', 'delegatedPermissions'],
        runsWhen: (args) => args.options.delegatedPermissions === undefined && args.options.applicationPermissions === undefined
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('appId', 'appObjectId', 'appName', 'applicationPermissions', 'delegatedPermissions');
    this.types.boolean.push('revokeAdminConsent');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeAppPermissions = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing permissions from application ${args.options.appId || args.options.appObjectId || args.options.appName}...`);
        }

        const appObject = await this.getAppObject(args.options);
        const servicePrincipals = await odata.getAllItems<ServicePrincipal>(`${this.resource}/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);

        const appPermissions: AppPermission[] = [];

        if (args.options.delegatedPermissions) {
          const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
          this.removePermissionsFromResourceArray(delegatedPermissions, appObject.requiredResourceAccess!);
        }

        if (args.options.applicationPermissions) {
          const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
          this.removePermissionsFromResourceArray(applicationPermissions, appObject.requiredResourceAccess!);
        }

        for (let i = 0; i < appObject.requiredResourceAccess!.length; i++) {
          if (appObject.requiredResourceAccess![i].resourceAccess?.length === 0) {
            appObject.requiredResourceAccess!.splice(i, 1);
          }
        }

        const removePermissionRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/applications/${appObject.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            requiredResourceAccess: appObject.requiredResourceAccess
          }
        };

        await request.patch(removePermissionRequestOptions);

        if (args.options.revokeAdminConsent) {
          const appServicePrincipal = servicePrincipals.find(sp => sp.appId === appObject.appId);
          await this.revokeAdminConsent(appServicePrincipal!, appPermissions, logger);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeAppPermissions();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the permissions from the specified application ${args.options.appId || args.options.appObjectId || args.options.appName}?` });

      if (result) {
        await removeAppPermissions();
      }
    }
  }

  private async getAppObject(options: Options): Promise<Application> {
    const selectProperties = '$select=id,appId,requiredResourceAccess';
    if (options.appObjectId) {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/applications/${options.appObjectId}?${selectProperties}`,
        headers: {
          'content-type': 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      return request.get<Application>(requestOptions);
    }

    const apps = options.appId
      ? await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=appId eq '${options.appId}'&${selectProperties}`)
      : await odata.getAllItems<Application>(`${this.resource}/v1.0/applications?$filter=displayName eq '${options.appName}'&${selectProperties}`);

    if (apps.length === 0) {
      throw `App with ${options.appId ? 'id' : 'name'} ${options.appId || options.appName} not found in Microsoft Entra ID`;
    }

    if (apps.length > 1) {
      const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', apps);
      return cli.handleMultipleResultsFound<Application>(`Multiple apps with name '${options.appName}' found.`, resultAsKeyValuePair);
    }

    return apps[0];
  }

  private async revokeAdminConsent(servicePrincipal: ServicePrincipal, appPermissions: AppPermission[], logger: Logger): Promise<void> {
    // Check if contains app permissions
    let appRoleAssignments: AppRoleAssignment[];
    let oAuth2RoleAssignments: OAuth2PermissionGrant[];

    if (appPermissions.some(perm => perm.resourceAccess.filter(acc => acc.type === ScopeType.Role).length > 0)) {
      // Retrieve app role assignments from service application
      appRoleAssignments = await odata.getAllItems<AppRoleAssignment>(`${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoleAssignments?$select=id,appRoleId,resourceId`);
    }

    if (appPermissions.filter(perm => perm.scope.length > 0).length > 0) {
      // Retrieve app role assignments from service application
      oAuth2RoleAssignments = await odata.getAllItems<OAuth2PermissionGrant>(`${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oAuth2PermissionGrants?$select=id,resourceId,scope`);
    }

    for await (const permission of appPermissions) {
      if (permission.scope.length > 0) {
        if (this.verbose) {
          await logger.logToStderr(`Revoking consent for delegated permission(s) with resourceId ${permission.resourceId} and scope(s) ${permission.scope.join(' ')}`);
        }

        const oAuth2RoleAssignment = oAuth2RoleAssignments!.find(y => y.resourceId === permission.resourceId);
        if (oAuth2RoleAssignment) {
          const scopes = oAuth2RoleAssignment?.scope?.split(' ');
          permission.scope.forEach(scope => {
            scopes!.splice(scopes!.indexOf(scope), 1);
          });

          oAuth2RoleAssignment!.scope = scopes!.join(' ');

          await this.revokeOAuth2Permission(oAuth2RoleAssignment!);
        }
      }

      for await (const access of permission.resourceAccess.filter(acc => acc.type === ScopeType.Role)) {
        if (this.verbose) {
          await logger.logToStderr(`Revoking consent for application permission with resourceId ${permission.resourceId} and appRoleId ${access.id}`);
        }

        const appRoleAssignmentToRemove = appRoleAssignments!.find(y => y.resourceId === permission.resourceId && y.appRoleId === access.id);
        if (appRoleAssignmentToRemove) {
          await this.revokeApplicationPermission(servicePrincipal.id!, appRoleAssignmentToRemove!.id!);
        }
      }
    }
  }

  private async revokeOAuth2Permission(oAuth2RoleAssignment: OAuth2PermissionGrant): Promise<void> {
    const revokeRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/oauth2PermissionGrants/${oAuth2RoleAssignment.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: oAuth2RoleAssignment
    };

    return request.patch<void>(revokeRequestOptions);
  }

  private async revokeApplicationPermission(servicePrincipalId: string, id: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments/${id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete<void>(requestOptions);
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

      let permission: AppRole | PermissionScope | undefined;

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

  private removePermissionsFromResourceArray(permissions: RequiredResourceAccess[], existingArray: RequiredResourceAccess[]): void {
    permissions.forEach(resolvedRequiredResource => {
      const requiredResource = existingArray?.find(api => api.resourceAppId === resolvedRequiredResource.resourceAppId);
      if (requiredResource) {
        resolvedRequiredResource.resourceAccess!.forEach(resolvedResourceAccess => {
          requiredResource.resourceAccess = requiredResource.resourceAccess!.filter(ra => ra.id !== resolvedResourceAccess.id);
        });
      }
    });
  }
}

export default new EntraAppPermissionRemoveCommand();