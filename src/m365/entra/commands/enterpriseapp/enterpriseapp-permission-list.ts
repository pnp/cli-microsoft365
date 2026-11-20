import { AppRoleAssignment, OAuth2PermissionGrant, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const allowedTypes = { delegated: 'delegated', application: 'application', all: 'all' } as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.guid().optional().alias('i'),
  appObjectId: z.guid().optional(),
  appName: z.string().optional().alias('n'),
  type: zod.coercedEnum(allowedTypes).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ApiPermission {
  resource: string;
  resourceId: string;
  permission: string;
  type: string;
}

class EntraEnterpriseAppPermissionListCommand extends GraphCommand {
  public get name(): string {
    return commands.ENTERPRISEAPP_PERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists the application and delegated permissions for a specified Entra enterprise application';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appName].filter(o => o !== undefined).length === 1, {
        error: 'Specify either appId, appObjectId, or appName',
        params: {
          customCode: 'optionSet',
          options: ['appId', 'appObjectId', 'appName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const type = args.options.type ?? 'all';
      const servicePrincipalId = await this.getServicePrincipalId(args.options, logger);

      if (this.verbose) {
        await logger.logToStderr(`Retrieving permissions granted to enterprise application ${servicePrincipalId}...`);
      }

      const [appRoleAssignments, oauth2PermissionGrants] = await Promise.all([
        type === 'delegated' ? Promise.resolve([]) : odata.getAllItems<AppRoleAssignment>(`${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/appRoleAssignments`),
        type === 'application' ? Promise.resolve([]) : odata.getAllItems<OAuth2PermissionGrant>(`${this.resource}/v1.0/servicePrincipals/${servicePrincipalId}/oauth2PermissionGrants`)
      ]);

      const resourceIds = [...new Set([
        ...appRoleAssignments.map(appRoleAssignment => appRoleAssignment.resourceId!),
        ...oauth2PermissionGrants.map(oauth2PermissionGrant => oauth2PermissionGrant.resourceId!)
      ])];
      const resources = await Promise.all(resourceIds.map(resourceId => this.getResourceServicePrincipal(resourceId, logger)));

      const permissions: ApiPermission[] = [];

      appRoleAssignments.forEach(appRoleAssignment => {
        const resource = resources.find(r => r.id === appRoleAssignment.resourceId)!;
        permissions.push({
          resource: resource.displayName!,
          resourceId: resource.appId!,
          permission: resource.appRoles!.find(appRole => appRole.id === appRoleAssignment.appRoleId)?.value ?? appRoleAssignment.appRoleId!,
          type: 'Application'
        });
      });

      oauth2PermissionGrants.forEach(oauth2PermissionGrant => {
        const resource = resources.find(r => r.id === oauth2PermissionGrant.resourceId)!;
        (oauth2PermissionGrant.scope ?? '')
          .split(' ')
          .filter(scope => scope !== '')
          .forEach(scope => {
            permissions.push({
              resource: resource.displayName!,
              resourceId: resource.appId!,
              permission: scope,
              type: 'Delegated'
            });
          });
      });

      await logger.log(permissions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getServicePrincipalId(options: Options, logger: Logger): Promise<string> {
    if (options.appObjectId) {
      return options.appObjectId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving enterprise application ${options.appId ?? options.appName}...`);
    }

    const servicePrincipal = options.appId
      ? await entraServicePrincipal.getServicePrincipalByAppId(options.appId, 'id')
      : await entraServicePrincipal.getServicePrincipalByAppName(options.appName!, 'id');

    return servicePrincipal.id!;
  }

  private async getResourceServicePrincipal(resourceId: string, logger: Logger): Promise<ServicePrincipal> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving resource service principal ${resourceId}...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${resourceId}?$select=id,appId,displayName,appRoles`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<ServicePrincipal>(requestOptions);
  }
}

export default new EntraEnterpriseAppPermissionListCommand();
