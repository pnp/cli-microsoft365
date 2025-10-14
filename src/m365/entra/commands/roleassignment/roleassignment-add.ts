import { UnifiedRoleAssignment } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { validation } from '../../../../utils/validation.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraApp } from '../../../../utils/entraApp.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { entraUser } from '../../../../utils/entraUser.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  roleDefinitionId: z.uuid().optional(),
  roleDefinitionName: z.string().optional(),
  principal: z.string().refine(principal => validation.isValidGuid(principal) || validation.isValidUserPrincipalName(principal) || validation.isValidMailNickname(principal), {
    error: e => `'${e.input}' is not a valid GUID, UPN or group mail nickname.`
  }),
  userId: z.uuid().optional(),
  userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  administrativeUnitId: z.uuid().optional(),
  administrativeUnitName: z.string().optional(),
  applicationId: z.uuid().optional(),
  applicationObjectId: z.uuid().optional(),
  applicationName: z.string().optional(),
  servicePrincipalId: z.uuid().optional(),
  servicePrincipalName: z.string().optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Assign a Entra ID role to a user and specify the scope for which the user has been granted access';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.roleDefinitionId, options.roleDefinitionName].filter(o => o !== undefined).length === 1, {
        error: 'Specify either roleDefinitionId or roleDefinitionName'
      })
      .refine(options => Object.values([
        options.userId, options.userName, options.administrativeUnitId, options.administrativeUnitName, options.applicationId, options.applicationObjectId, options.applicationName,
        options.servicePrincipalId, options.servicePrincipalName, options.groupId, options.groupName]).filter(v => typeof v !== 'undefined').length < 2, {
        message: 'Provide value for only one of the following parameters: userId, userName, administrativeUnitId, administrativeUnitName, applicationId, applicationObjectId, applicationName, servicePrincipalId, servicePrincipalName, groupId or groupName'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      const data: UnifiedRoleAssignment = {
        roleDefinitionId: roleDefinitionId,
        principalId: await this.getPrincipalId(args.options, logger),
        directoryScopeId: await this.getDirectoryScopeId(args.options)
      };

      const requestOptions: any = {
        url: `${this.resource}/v1.0/roleManagement/directory/roleAssignments`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: data
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<string> {
    if (options.roleDefinitionId) {
      return options.roleDefinitionId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving role definition by its name '${options.roleDefinitionName}'`);
    }

    const role = await roleDefinition.getRoleDefinitionByDisplayName(options.roleDefinitionName!);
    return role.id!;
  }

  private async getPrincipalId(options: Options, logger: Logger): Promise<string> {
    if (validation.isValidGuid(options.principal)) {
      return options.principal;
    }

    if (validation.isValidUserPrincipalName(options.principal)) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user id by UPN '${options.principal}'`);
      }
      return await entraUser.getUserIdByUpn(options.principal);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving group id by mail nickname '${options.principal}'`);
    }
    return await entraGroup.getGroupIdByMailNickname(options.principal);
  }

  private async getDirectoryScopeId(options: Options): Promise<string> {
    let prefix = '/';
    let resourceId: string | undefined = '';

    if (options.userId || options.userName) {
      resourceId = options.userId || await entraUser.getUserIdByUpn(options.userName!);
    }
    else if (options.administrativeUnitId || options.administrativeUnitName) {
      prefix = '/administrativeUnits/';
      resourceId = options.administrativeUnitId || (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(options.administrativeUnitName!, "id"))!.id;
    }
    else if (options.applicationId || options.applicationObjectId || options.applicationName) {
      resourceId = options.applicationObjectId;
      if (options.applicationId) {
        resourceId = (await entraApp.getAppRegistrationByAppId(options.applicationId!, ["id"])).id;
      }
      else if (options.applicationName) {
        resourceId = (await entraApp.getAppRegistrationByAppName(options.applicationName!, ["id"])).id;
      }
    }
    else if (options.servicePrincipalId || options.servicePrincipalName) {
      resourceId = options.servicePrincipalId || (await entraServicePrincipal.getServicePrincipalByAppName(options.servicePrincipalName!, "id")).id;
    }
    else if (options.groupId || options.groupName) {
      resourceId = options.groupId || (await entraGroup.getGroupIdByDisplayName(options.groupName!));
    }

    return `${prefix}${resourceId}`;
  }
}

export default new EntraRoleAssignmentAddCommand();
