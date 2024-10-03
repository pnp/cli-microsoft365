import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraServicePrincipal } from '../../../../utils/entraServicePrincipal.js';

const options = globalOptionsZod
  .extend({
    roleDefinitionId: z.string().optional(),
    roleDefinitionName: z.string().optional(),
    principalId: z.string().optional(),
    principalName: z.string().optional(),
    scopeUserId: z.string().optional(),
    scopeUserName: z.string().optional(),
    scopeGroupId: z.string().optional(),
    scopeGroupName: z.string().optional(),
    scopeAdministrativeUnitId: z.string().optional(),
    scopeAdministrativeUnitName: z.string().optional(),
    scopeTenant: z.boolean().optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExchangeRoleAssignmentRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROLE_ASSIGNMENT_ADD;
  }

  public get description(): string {
    return `Grant permissions to an application that's accessing data in Exchange Online and specify which mailboxes an app can access.`;
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.roleDefinitionId !== !options.roleDefinitionName, {
        message: 'Specify either roleDefinitionId or roleDefinitionName, but not both'
      })
      .refine(options => options.roleDefinitionId || options.roleDefinitionName, {
        message: 'Specify either roleDefinitionId or roleDefinitionName'
      })
      .refine(options => (!options.roleDefinitionId && !options.roleDefinitionName) || options.roleDefinitionName || (options.roleDefinitionId && validation.isValidGuid(options.roleDefinitionId)), options => ({
        message: `The '${options.roleDefinitionId}' must be a valid GUID`,
        path: ['roleDefinitionId']
      }))
      .refine(options => !options.principalId !== !options.principalName, {
        message: 'Specify either principalId or principalName, but not both'
      })
      .refine(options => options.principalId || options.principalName, {
        message: 'Specify either principalId or principalName'
      })
      .refine(options => (!options.principalId && !options.principalName) || options.principalName || (options.principalId && validation.isValidGuid(options.principalId)), options => ({
        message: `The '${options.principalId}' must be a valid GUID`,
        path: ['principalId']
      }))
      .refine(options => Object.values([options.scopeTenant, options.scopeUserId, options.scopeUserName, options.scopeGroupId, options.scopeGroupName, options.scopeAdministrativeUnitId, options.scopeAdministrativeUnitName]).filter(v => typeof v !== 'undefined').length === 1, {
        message: 'Specify either scopeTenant, scopeUserId, scopeUserName, scopeGroupId, scopeGroupName, scopeAdministrativeUnitId, or scopeAdministrativeUnitName, but not multiple'
      })
      .refine(options => (!options.scopeTenant && !options.scopeUserId && !options.scopeUserName && !options.scopeGroupId && !options.scopeGroupName && !options.scopeAdministrativeUnitId && !options.scopeAdministrativeUnitName)
        || options.scopeTenant || options.scopeUserName || options.scopeGroupId || options.scopeGroupName || options.scopeAdministrativeUnitId || options.scopeAdministrativeUnitName ||
        (options.scopeUserId && validation.isValidGuid(options.scopeUserId)), options => ({
        message: `The '${options.scopeUserId}' must be a valid GUID`,
        path: ['scopeUserId']
      }))
      .refine(options => (!options.scopeTenant && !options.scopeUserId && !options.scopeUserName && !options.scopeGroupId && !options.scopeGroupName && !options.scopeAdministrativeUnitId && !options.scopeAdministrativeUnitName)
        || options.scopeTenant || options.scopeUserId || options.scopeGroupId || options.scopeGroupName || options.scopeAdministrativeUnitId || options.scopeAdministrativeUnitName ||
        (options.scopeUserName && validation.isValidUserPrincipalName(options.scopeUserName)), options => ({
        message: `The '${options.scopeUserName}' must be a valid UPN`,
        path: ['scopeUserName']
      }))
      .refine(options => (!options.scopeTenant && !options.scopeUserId && !options.scopeUserName && !options.scopeGroupId && !options.scopeGroupName && !options.scopeAdministrativeUnitId && !options.scopeAdministrativeUnitName)
        || options.scopeTenant || options.scopeUserId || options.scopeUserName || options.scopeGroupName || options.scopeAdministrativeUnitId || options.scopeAdministrativeUnitName ||
        (options.scopeGroupId && validation.isValidGuid(options.scopeGroupId)), options => ({
        message: `The '${options.scopeGroupId}' must be a valid GUID`,
        path: ['scopeGroupId']
      }))
      .refine(options => (!options.scopeTenant && !options.scopeUserId && !options.scopeUserName && !options.scopeGroupId && !options.scopeGroupName && !options.scopeAdministrativeUnitId && !options.scopeAdministrativeUnitName)
        || options.scopeTenant || options.scopeUserId || options.scopeUserName || options.scopeGroupId || options.scopeGroupName || options.scopeAdministrativeUnitName ||
        (options.scopeAdministrativeUnitId && validation.isValidGuid(options.scopeAdministrativeUnitId)), options => ({
        message: `The '${options.scopeAdministrativeUnitId}' must be a valid GUID`,
        path: ['scopeAdministrativeUnitId']
      }));
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);

      const data = {
        roleDefinitionId: roleDefinitionId,
        principalId: await this.getPrincipalId(args.options),
        directoryScopeId: await this.getDirectoryScopeId(args.options)
      };

      const requestOptions: any = {
        url: `${this.resource}/beta/roleManagement/exchange/roleAssignments`,
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

    const role = await roleDefinition.getExchangeRoleDefinitionByDisplayName(options.roleDefinitionName!);
    return role.id!;
  }

  private async getPrincipalId(options: Options): Promise<string> {
    let principalId = '';
    if (options.principalId) {
      principalId = options.principalId;
    }
    else {
      principalId = await entraServicePrincipal.getServicePrincipalIdFromAppName(options.principalName!);
    }

    return `/ServicePrincipals/${principalId}`;
  }

  private async getDirectoryScopeId(options: Options): Promise<string> {
    let prefix = '/';
    let resourceId = '';
    if (options.scopeUserId) {
      prefix = '/users/';
      resourceId = options.scopeUserId;
    }
    else if (options.scopeUserName) {
      prefix = '/users/';
      resourceId = await entraUser.getUserIdByUpn(options.scopeUserName);
    }
    else if (options.scopeGroupId) {
      prefix = '/groups/';
      resourceId = options.scopeGroupId;
    }
    else if (options.scopeGroupName) {
      prefix = '/groups/';
      resourceId = await entraGroup.getGroupIdByDisplayName(options.scopeGroupName);
    }
    else if (options.scopeAdministrativeUnitId) {
      prefix = '/administrativeUnits/';
      resourceId = options.scopeAdministrativeUnitId;
    }
    else if (options.scopeAdministrativeUnitName) {
      prefix = '/administrativeUnits/';
      const administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(options.scopeAdministrativeUnitName);
      resourceId = administrativeUnit.id!;
    }

    return `${prefix}${resourceId}`;
  }
}

export default new ExchangeRoleAssignmentRoleAssignmentAddCommand();