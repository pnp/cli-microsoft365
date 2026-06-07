import os from 'os';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { ServicePrincipal } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';

interface AppRole {
  objectId: string;
  value: string;
  resourceId: string;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional(),
  appObjectId: z.uuid().optional(),
  appDisplayName: z.string().optional(),
  resource: z.string().alias('r'),
  scopes: z.string().transform((value) => value.split(',').map(String)).alias('s')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds service principal permissions also known as scopes and app role assignments for specified Microsoft Entra application registration';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let objectId: string;
    let queryFilter: string;
    if (args.options.appId) {
      queryFilter = `$filter=appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
    }
    else if (args.options.appObjectId) {
      queryFilter = `$filter=id eq '${formatting.encodeQueryParameter(args.options.appObjectId)}'`;
    }
    else {
      queryFilter = `$filter=displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
    }

    const getServicePrinciplesRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals?${queryFilter}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const servicePrincipalResult = await request.get<{ value: ServicePrincipal[] }>(getServicePrinciplesRequestOptions);

      if (servicePrincipalResult.value.length === 0) {
        throw `The specified service principal doesn't exist`;
      }

      if (servicePrincipalResult.value.length > 1) {
        const resultAsKeyValuePair = formatting.convertArrayToHashTable('id', servicePrincipalResult.value);
        const result = await cli.handleMultipleResultsFound<ServicePrincipal>(`Multiple service principal found.`, resultAsKeyValuePair);
        objectId = result.id!;
      }
      else {
        objectId = servicePrincipalResult.value[0].id!;
      }


      let resource: string = formatting.encodeQueryParameter(args.options.resource);

      // try resolve aliases that the user might enter since these are seen in the Azure portal
      switch (args.options.resource.toLocaleLowerCase()) {
        case 'sharepoint':
          resource = 'Office 365 SharePoint Online';
          break;
        case 'intune':
          resource = 'Microsoft Intune API';
          break;
        case 'exchange':
          resource = 'Office 365 Exchange Online';
          break;
      }

      // will perform resource name, appId or objectId search
      let filter: string = `$filter=(displayName eq '${resource}' or startswith(displayName,'${resource}'))`;

      if (validation.isValidGuid(resource)) {
        filter += ` or appId eq '${resource}' or id eq '${resource}'`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/servicePrincipals?${filter}`,
        headers: {
          'accept': 'application/json'
        },
        responseType: 'json'
      };

      const res = await request.get<{ value: ServicePrincipal[] }>(requestOptions);

      const appRoles: AppRole[] = [];

      // flatten the app roles found
      const appRolesFound: AppRole[] = [];
      for (const servicePrincipal of res.value) {
        for (const role of servicePrincipal.appRoles!) {
          appRolesFound.push({
            resourceId: servicePrincipal.id!,
            objectId: role.id!,
            value: role.value!
          });
        }
      }

      if (!appRolesFound.length) {
        throw `The resource '${args.options.resource}' does not have any application permissions available.`;
      }

      // search for match between the found app roles and the specified scopes option value
      for (const scope of args.options.scopes) {
        const existingRoles = appRolesFound.filter((role: AppRole) => {
          return role.value.toLocaleLowerCase() === scope.toLocaleLowerCase().trim();
        });

        if (!existingRoles.length) {
          // the role specified in the scopes option does not belong to the found service principles
          // throw an error and show list with available roles (scopes)
          let availableRoles: string = '';
          appRolesFound.map((r: AppRole) => availableRoles += `${os.EOL}${r.value}`);

          throw `The scope value '${scope}' you have specified does not exist for ${args.options.resource}. ${os.EOL}Available scopes (application permissions) are: ${availableRoles}`;
        }

        appRoles.push(existingRoles[0]);
      }

      const tasks: Promise<any>[] = [];

      for (const appRole of appRoles) {
        tasks.push(this.addRoleToServicePrincipal(objectId, appRole));
      }

      const rolesAddedResponse = await Promise.all(tasks);

      if (args.options.output && args.options.output.toLowerCase() === 'json') {
        await logger.log(rolesAddedResponse);
      }
      else {
        await logger.log(rolesAddedResponse.map((result: any) => ({
          objectId: result.id,
          principalDisplayName: result.principalDisplayName,
          resourceDisplayName: result.resourceDisplayName
        })));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleToServicePrincipal(objectId: string, appRole: AppRole): Promise<any> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${objectId}/appRoleAssignments`,
      headers: {
        'Content-Type': 'application/json'
      },
      responseType: 'json',
      data: {
        appRoleId: appRole.objectId,
        principalId: objectId,
        resourceId: appRole.resourceId
      }
    };

    return request.post(requestOptions);
  }
}

export default new EntraAppRoleAssignmentAddCommand();