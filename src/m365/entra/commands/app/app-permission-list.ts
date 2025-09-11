import { AppRole, Application, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { z } from 'zod';
import { Logger } from "../../../../cli/Logger.js";
import { globalOptionsZod } from "../../../../Command.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { entraApp } from "../../../../utils/entraApp.js";
import { zod } from "../../../../utils/zod.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";

const allowedTypes = ['delegated', 'application', 'all'] as const;

const options = globalOptionsZod
  .extend({
    appId: zod.alias('i', z.string().uuid().optional()),
    appName: zod.alias('n', z.string().optional()),
    appObjectId: z.string().uuid().optional(),
    type: z.enum(allowedTypes).optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ApiPermission {
  resource: string;
  resourceId?: string;
  permission: string;
  type: string;
}

interface ServicePrincipalInfo {
  appId?: string;
  id?: string;
}

class EntraAppPermissionListCommand extends GraphCommand {
  public get name(): string {
    return commands.APP_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists the application and delegated permissions for a specified Entra Application Registration';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appName, options.appObjectId].filter(Boolean).length === 1, {
        message: 'Specify either appId, appName, or appObjectId'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appObjectId = await this.getAppObjectId(args.options, logger);
      const type = args.options.type ?? 'all';
      const permissions = await this.getAppRegPermissions(appObjectId, type, logger);
      await logger.log(permissions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(options: Options, logger: Logger): Promise<string> {
    if (options.appObjectId) {
      return options.appObjectId;
    }

    const { appId, appName } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving information about Microsoft Entra app ${appId ? appId : appName}...`);
    }

    if (appId) {
      const app = await entraApp.getAppRegistrationByAppId(appId, ["id"]);
      return app.id!;
    }
    else {
      const app = await entraApp.getAppRegistrationByAppName(appName!, ["id"]);
      return app.id!;
    }
  }

  private async getAppRegPermissions(appObjectId: string, permissionType: string, logger: Logger): Promise<ApiPermission[]> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications/${appObjectId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const application = await request.get<Application>(requestOptions);
    const requiredResourceAccess = application.requiredResourceAccess as RequiredResourceAccess[];

    if (requiredResourceAccess.length === 0) {
      return [];
    }

    const servicePrincipalsToResolve: ServicePrincipalInfo[] =
      requiredResourceAccess.map(resourceAccess => {
        return {
          appId: resourceAccess.resourceAppId as string
        };
      });
    const servicePrincipals = await Promise
      .all(servicePrincipalsToResolve.map(servicePrincipalInfo =>
        this.getServicePrincipal(servicePrincipalInfo, permissionType, logger) as ServicePrincipal));

    const apiPermissions: ApiPermission[] = [];
    requiredResourceAccess.forEach(requiredResourceAccess => {
      const servicePrincipal = servicePrincipals
        .find(servicePrincipal => servicePrincipal?.appId === requiredResourceAccess.resourceAppId as string);
      const resourceName = servicePrincipal?.displayName as string ?? requiredResourceAccess.resourceAppId as string;
      (requiredResourceAccess.resourceAccess as ResourceAccess[]).forEach(permission => {
        if (permissionType === 'application' && permission.type === 'Scope') { return; }
        if (permissionType === 'delegated' && permission.type === 'Role') { return; }

        apiPermissions.push({
          resource: resourceName,
          resourceId: requiredResourceAccess.resourceAppId,
          permission: this.getPermissionName(permission.id as string, permission.type as string, servicePrincipal),
          type: permission.type === 'Role' ? 'Application' : 'Delegated'
        });
      });
    });

    return apiPermissions;
  }

  private getPermissionName(permissionId: string, permissionType: string, servicePrincipal: ServicePrincipal | undefined): string {
    if (!servicePrincipal) {
      return permissionId;
    }

    if (permissionType === 'Role') {
      return (servicePrincipal.appRoles as AppRole[])
        .find(appRole => appRole.id === permissionId)?.value as string ?? permissionId;
    }

    // permissionType === 'Scope'
    return (servicePrincipal.oauth2PermissionScopes as PermissionScope[])
      .find(permissionScope => permissionScope.id === permissionId)?.value as string ?? permissionId;
  }

  private async getServicePrincipal(servicePrincipalInfo: ServicePrincipalInfo, permissionType: string, logger: Logger): Promise<ServicePrincipal | null> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving service principal ${servicePrincipalInfo.appId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals?$filter=appId eq '${servicePrincipalInfo.appId}'&$select=appId,id,displayName`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: ServicePrincipal[] }>(requestOptions);

    if (servicePrincipalInfo.appId && response.value.length === 0) {
      return null;
    }

    const servicePrincipal = response.value[0];

    if (this.verbose) {
      await logger.logToStderr(`Retrieving permissions for service principal ${servicePrincipal.id}...`);
    }

    const oauth2PermissionScopesRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/oauth2PermissionScopes`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const appRolesRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${servicePrincipal.id}/appRoles`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    let permissions: any;
    if (permissionType === 'all' || permissionType === 'delegated') {
      permissions = await request.get<{ value: PermissionScope[] }>(oauth2PermissionScopesRequestOptions);
      servicePrincipal.oauth2PermissionScopes = permissions.value as PermissionScope[];
    }

    if (permissionType === 'all' || permissionType === 'application') {
      permissions = await request.get<{ value: AppRole[] }>(appRolesRequestOptions);
      servicePrincipal.appRoles = permissions.value as AppRole[];
    }

    return servicePrincipal;
  }
}

export default new EntraAppPermissionListCommand();