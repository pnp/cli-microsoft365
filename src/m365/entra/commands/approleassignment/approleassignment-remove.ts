import os from 'os';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { AppRole, AppRoleAssignment, ServicePrincipal } from '@microsoft/microsoft-graph-types';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  appId: z.uuid().optional(),
  appObjectId: z.uuid().optional(),
  appDisplayName: z.string().optional(),
  resource: z.string().alias('r'),
  scopes: z.string().transform((value) => value.split(',').map(String)).alias('s'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAppRoleAssignmentRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Deletes an app role assignment for the specified Entra Application Registration';
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
    const removeAppRoleAssignment = async (): Promise<void> => {
      let sp: ServicePrincipal;
      // get the service principal associated with the appId
      let spMatchQuery: string;
      if (args.options.appId) {
        spMatchQuery = `appId eq '${formatting.encodeQueryParameter(args.options.appId)}'`;
      }
      else if (args.options.appObjectId) {
        spMatchQuery = `id eq '${formatting.encodeQueryParameter(args.options.appObjectId)}'`;
      }
      else {
        spMatchQuery = `displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
      }

      try {
        let resp = await this.getServicePrincipalForApp(spMatchQuery);

        if (!resp.value.length) {
          throw 'app registration not found';
        }

        sp = resp.value[0];
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

        resp = await request.get<{ value: ServicePrincipal[] }>(requestOptions);

        if (!resp.value.length) {
          throw 'Resource not found';
        }

        const appRolesToBeDeleted: AppRole[] = [];
        const appRolesFound: AppRole[] = resp.value[0].appRoles!;

        if (!appRolesFound.length) {
          throw `The resource '${args.options.resource}' does not have any application permissions available.`;
        }

        for (const scope of args.options.scopes) {
          const existingRoles = appRolesFound.filter((role: AppRole) => {
            return role.value!.toLocaleLowerCase() === scope.toLocaleLowerCase().trim();
          });

          if (!existingRoles.length) {
            // the role specified in the scopes option does not belong to the found service principles
            // throw an error and show list with available roles (scopes)
            let availableRoles: string = '';
            appRolesFound.map((r: AppRole) => availableRoles += `${os.EOL}${r.value}`);

            throw `The scope value '${scope}' you have specified does not exist for ${args.options.resource}. ${os.EOL}Available scopes (application permissions) are: ${availableRoles}`;
          }

          appRolesToBeDeleted.push(existingRoles[0]);
        }
        const tasks: Promise<any>[] = [];

        for (const appRole of appRolesToBeDeleted) {
          const appRoleAssignment = sp.appRoleAssignments!.filter((role: AppRoleAssignment) => role.appRoleId === appRole.id);
          if (!appRoleAssignment.length) {
            throw 'App role assignment not found';
          }
          tasks.push(this.removeAppRoleAssignmentForServicePrincipal(sp.id!, appRoleAssignment[0].id!));
        }

        await Promise.all(tasks);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeAppRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the appRoleAssignment with scope(s) ${args.options.scopes} for resource ${args.options.resource}?` });

      if (result) {
        await removeAppRoleAssignment();
      }
    }
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

  private async removeAppRoleAssignmentForServicePrincipal(spId: string, appRoleAssignmentId: string): Promise<ServicePrincipal> {
    const spRequestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/servicePrincipals/${spId}/appRoleAssignments/${appRoleAssignmentId}`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.delete(spRequestOptions);
  }
}

export default new EntraAppRoleAssignmentRemoveCommand();