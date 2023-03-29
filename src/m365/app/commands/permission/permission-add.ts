import { Application, AppRole, PermissionScope, RequiredResourceAccess, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { Cli } from '../../../../cli/Cli';
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
        grantAdminConsent: typeof !!args.options.grantAdminConsent
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
      const servicePrincipals = await odata.getAllItems<ServicePrincipal>(`${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,appRoles,id,oauth2PermissionScopes,servicePrincipalNames`);
      const appObject = await this.getAppObjectId();

      let delegatedPermissions: RequiredResourceAccess[] = [], applicationPermissions: RequiredResourceAccess[] = [];

      if (args.options.delegatedPermission) {
        delegatedPermissions = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.delegatedPermission, 'Scope', logger);
        this.addPermissionsToResourceArray(delegatedPermissions, appObject.requiredResourceAccess!);
      }
      if (args.options.applicationPermission) {
        applicationPermissions = this.getRequiredResourceAccessForApis(servicePrincipals, args.options.applicationPermission, 'Role', logger);
        this.addPermissionsToResourceArray(applicationPermissions, appObject.requiredResourceAccess!);
      }

      /*const addPermissionsRequestOptions: CliRequestOptions = {
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
*/
      if (args.options.grantAdminConsent) {
        await this.grantAdminConsent(appObject, logger, delegatedPermissions, applicationPermissions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async grantAdminConsent(appObject: Application, logger: Logger, delegatedPermissions: RequiredResourceAccess[], applicationPermissions: RequiredResourceAccess[]): Promise<void> {
    logger.log(appObject);
    logger.log(delegatedPermissions);
    logger.log(applicationPermissions);
  }

  private async getAppObjectId(): Promise<Application> {
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
        Cli.log(resolvedRequiredResource);
        Cli.log(requiredResource);
        requiredResource.resourceAccess!.push(...resolvedRequiredResource.resourceAccess!);
      }
      else {
        existingArray.push(resolvedRequiredResource);
      }
    });
  }

  private getRequiredResourceAccessForApis(servicePrincipals: ServicePrincipal[], apis: string, scopeType: string, logger: Logger): RequiredResourceAccess[] {
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

      if (scopeType === 'Scope') {
        permission = servicePrincipal.oauth2PermissionScopes!.find(scope => scope.value === permissionName);
      }
      else if (scopeType === 'Role') {
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

      resolvedApi.resourceAccess!.push({
        id: permission.id,
        type: scopeType
      });
    });
    return resolvedApis;
  }
}

module.exports = new AppPermissionAddCommand();