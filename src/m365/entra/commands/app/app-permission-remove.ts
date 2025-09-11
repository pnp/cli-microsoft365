import { AppRole, AppRoleAssignment, Application, OAuth2PermissionGrant, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { cli } from "../../../../cli/cli.js";
import { entraApp } from "../../../../utils/entraApp.js";
import { odata } from "../../../../utils/odata.js";
import { validation } from "../../../../utils/validation.js";
import { zod } from "../../../../utils/zod.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";

const options = globalOptionsZod
  .extend({
    appId: z.string().uuid().optional(),
    appObjectId: z.string().uuid().optional(),
    appName: z.string().optional(),
    applicationPermissions: z.string().optional(),
    delegatedPermissions: z.string().optional(),
    revokeAdminConsent: z.boolean().optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(
        options => [options.appId, options.appObjectId, options.appName].filter(Boolean).length === 1,
        'Specify either appId, appObjectId, or appName'
      )
      .refine(
        options => options.applicationPermissions || options.delegatedPermissions,
        'Specify either applicationPermissions or delegatedPermissions'
      )
      .refine(
        options => !options.delegatedPermissions || !Array.isArray(validation.isValidPermission(options.delegatedPermissions)),
        options => ({
          message: `Delegated permission(s) ${(validation.isValidPermission(options.delegatedPermissions!) as string[]).join(', ')} are not fully-qualified`
        })
      )
      .refine(
        options => !options.applicationPermissions || !Array.isArray(validation.isValidPermission(options.applicationPermissions)),
        options => ({
          message: `Application permission(s) ${(validation.isValidPermission(options.applicationPermissions!) as string[]).join(', ')} are not fully-qualified`
        })
      );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeAppPermissions = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing permissions from application ${args.options.appId || args.options.appObjectId || args.options.appName}...`);
        }

        const entraApp = await this.getEntraApp(args.options, logger);
        const servicePrincipals = await odata.getAllItems<ServicePrincipal>(`${this.resource}/v1.0/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);

        const appPermissions: AppPermission[] = [];

        if (args.options.delegatedPermissions) {
          const delegatedPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermissions, ScopeType.Scope, appPermissions, logger);
          this.removePermissionsFromResourceArray(delegatedPermissions, entraApp.requiredResourceAccess!);
        }

        if (args.options.applicationPermissions) {
          const applicationPermissions = await this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermissions, ScopeType.Role, appPermissions, logger);
          this.removePermissionsFromResourceArray(applicationPermissions, entraApp.requiredResourceAccess!);
        }

        for (let i = 0; i < entraApp.requiredResourceAccess!.length; i++) {
          if (entraApp.requiredResourceAccess![i].resourceAccess?.length === 0) {
            entraApp.requiredResourceAccess!.splice(i, 1);
          }
        }

        const removePermissionRequestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/applications/${entraApp.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            requiredResourceAccess: entraApp.requiredResourceAccess
          }
        };

        await request.patch(removePermissionRequestOptions);

        if (args.options.revokeAdminConsent) {
          const appServicePrincipal = servicePrincipals.find(sp => sp.appId === entraApp.appId);

          if (appServicePrincipal) {
            await this.revokeAdminConsent(appServicePrincipal, appPermissions, logger);
          }
          else {
            if (this.debug) {
              await logger.logToStderr(`No service principal found for the appId: ${entraApp.appId}. Skipping revoking admin consent.`);
            }
          }
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

  private async getEntraApp(options: Options, logger: Logger): Promise<Application> {
    const { appObjectId, appId, appName } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appObjectId ? appObjectId : (appId ? appId : appName)}...`);
    }

    if (appObjectId) {
      return await entraApp.getAppRegistrationByObjectId(appObjectId, ['id', 'appId', 'requiredResourceAccess']);
    }
    else if (appId) {
      return await entraApp.getAppRegistrationByAppId(appId, ['id', 'appId', 'requiredResourceAccess']);
    }
    else {
      return await entraApp.getAppRegistrationByAppName(appName!, ['id', 'appId', 'requiredResourceAccess']);
    }
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