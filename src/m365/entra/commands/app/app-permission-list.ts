import { AppRole, Application, PermissionScope, RequiredResourceAccess, ResourceAccess, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import GlobalOptions from "../../../../GlobalOptions.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";
import request, { CliRequestOptions } from "../../../../request.js";
import { Logger } from "../../../../cli/Logger.js";
import { validation } from "../../../../utils/validation.js";
import { formatting } from "../../../../utils/formatting.js";


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  appObjectId?: string;
  type?: string;
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
  private allowedTypes: string[] = ['delegated', 'application', 'all'];

  public get name(): string {
    return commands.APP_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists the application and delegated permissions for a specified Entra Application Registration';
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
        appObjectId: typeof args.options.appObjectId !== 'undefined',
        type: typeof args.options.type !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --appId [appId]' },
      { option: '--appObjectId [appObjectId]' },
      { option: '--type [type]', autocomplete: this.allowedTypes }
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

        if (args.options.type && this.allowedTypes.indexOf(args.options.type.toLowerCase()) === -1) {
          return `${args.options.type} is not a valid type. Allowed types are ${this.allowedTypes.join(', ')}`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appObjectId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const appObjectId = await this.getAppObjectId(args.options);
      const type = args.options.type ?? 'all';
      const permissions = await this.getAppRegPermissions(appObjectId, type, logger);
      await logger.log(permissions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppObjectId(options: Options): Promise<string> {
    if (options.appObjectId) {
      return options.appObjectId;
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(options.appId!)}'&$select=id`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { id: string }[] }>(requestOptions);

    if (res.value.length === 0) {
      throw `No Azure AD application registration with ID ${options.appId} found`;
    }

    return res.value[0].id;
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