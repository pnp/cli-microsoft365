import { AppRole, AppRoleAssignment, ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional().alias('i'),
  appDisplayName: z.string().optional().alias('n'),
  appObjectId: z.uuid().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleAssignmentListCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists app role assignments for the specified application registration';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.appId, options.appObjectId, options.appDisplayName].filter(o => o !== undefined).length === 1, {
        error: 'Specify either appId, appObjectId, or appDisplayName',
        params: {
          customCode: 'optionSet',
          options: ['appId', 'appObjectId', 'appDisplayName']
        }
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['resourceDisplayName', 'roleName'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spAppRoleAssignments = await this.getAppRoleAssignments(args.options);
      // the role assignment has an appRoleId but no name. To get the name,
      // we need to get all the roles from the resource. the resource is
      // a service principal. Multiple roles may have same resource id.
      const resourceIds = spAppRoleAssignments.map((item: AppRoleAssignment) => item.resourceId);

      const tasks: Promise<ServicePrincipal>[] = [];
      for (let i: number = 0; i < resourceIds.length; i++) {
        tasks.push(this.getServicePrincipal(resourceIds[i]!));
      }

      const resources = await Promise.all(tasks);

      // loop through all appRoleAssignments for the servicePrincipal
      // and lookup the appRole.Id in the resources[resourceId].appRoles array...
      const results: any[] = [];
      spAppRoleAssignments.map((appRoleAssignment: AppRoleAssignment) => {
        const resource: ServicePrincipal | undefined = resources.find((r: any) => r.id === appRoleAssignment.resourceId);

        if (resource) {
          const appRole: AppRole | undefined = resource.appRoles!.find((r: any) => r.id === appRoleAssignment.appRoleId);

          if (appRole) {
            results.push({
              appRoleId: appRoleAssignment.appRoleId,
              resourceDisplayName: appRoleAssignment.resourceDisplayName,
              resourceId: appRoleAssignment.resourceId,
              roleId: appRole.id,
              roleName: appRole.value,
              created: appRoleAssignment.createdDateTime,
              deleted: appRoleAssignment.deletedDateTime
            });
          }
        }
      });

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getAppRoleAssignments(argOptions: Options): Promise<AppRoleAssignment[]> {
    if (argOptions.appObjectId) {
      const spAppRoleAssignments = await this.getSPAppRoleAssignments(argOptions.appObjectId);

      if (!spAppRoleAssignments.value.length) {
        throw 'no app role assignments found';
      }

      return spAppRoleAssignments.value;
    }
    else {
      const spMatchQuery: string = argOptions.appId
        ? `appId eq '${formatting.encodeQueryParameter(argOptions.appId)}'`
        : `displayName eq '${formatting.encodeQueryParameter(argOptions.appDisplayName as string)}'`;

      const resp = await this.getServicePrincipalForApp(spMatchQuery);
      if (!resp.value.length) {
        throw 'app registration not found';
      }

      return resp.value[0].appRoleAssignments!;
    }
  }

  private async getSPAppRoleAssignments(spId: string): Promise<{ value: AppRoleAssignment[] }> {
    const spRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${spId}/appRoleAssignments`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<{ value: AppRoleAssignment[] }>(spRequestOptions);
  }

  private async getServicePrincipalForApp(filterParam: string): Promise<{ value: ServicePrincipal[] }> {
    const spRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals?$expand=appRoleAssignments&$filter=${filterParam}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<{ value: ServicePrincipal[] }>(spRequestOptions);
  }

  private async getServicePrincipal(spId: string): Promise<ServicePrincipal> {
    const spRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${spId}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request.get<ServicePrincipal>(spRequestOptions);
  }
}

export default new EntraAppRoleAssignmentListCommand();