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
import { customAppScope } from '../../../../utils/customAppScope.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  roleDefinitionId: z.string().optional(),
  roleDefinitionName: z.string().optional(),
  principalId: z.string().optional(),
  principalName: z.string().optional(),
  scope: z.enum(['tenant', 'user', 'group', 'administrativeUnit', 'custom']).alias('s'),
  userId: z.string().optional(),
  userName: z.string().optional(),
  groupId: z.string().optional(),
  groupName: z.string().optional(),
  administrativeUnitId: z.string().optional(),
  administrativeUnitName: z.string().optional(),
  customAppScopeId: z.string().optional(),
  customAppScopeName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExoAppRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return `Grant permissions to an application that's accessing data in Exchange Online and specify which mailboxes an app can access.`;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !options.roleDefinitionId !== !options.roleDefinitionName, {
        error: 'Specify either roleDefinitionId or roleDefinitionName, but not both'
      })
      .refine(options => options.roleDefinitionId || options.roleDefinitionName, {
        error: 'Specify either roleDefinitionId or roleDefinitionName'
      })
      .refine(options => (!options.roleDefinitionId && !options.roleDefinitionName) || options.roleDefinitionName || (options.roleDefinitionId && validation.isValidGuid(options.roleDefinitionId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['roleDefinitionId']
      })
      .refine(options => !options.principalId !== !options.principalName, {
        error: 'Specify either principalId or principalName, but not both'
      })
      .refine(options => options.principalId || options.principalName, {
        error: 'Specify either principalId or principalName'
      })
      .refine(options => (!options.principalId && !options.principalName) || options.principalName || (options.principalId && validation.isValidGuid(options.principalId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['principalId']
      })
      .refine(options => options.scope !== 'tenant' || Object.values([options.userId, options.userName, options.groupId, options.groupName, options.administrativeUnitId, options.administrativeUnitName, options.customAppScopeId, options.customAppScopeName]).filter(v => typeof v !== 'undefined').length === 0, {
        message: "When the scope is set to 'tenant' then do not specify neither userId, userName, groupId, groupName, administrativeUnitId, administrativeUnitName, customAppScopeId nor customAppScopeName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'user' || Object.values([options.groupId, options.groupName, options.administrativeUnitId, options.administrativeUnitName, options.customAppScopeId, options.customAppScopeName]).filter(v => typeof v !== 'undefined').length === 0, {
        message: "When the scope is set to 'user' then do not specify groupId, groupName, administrativeUnitId, administrativeUnitName, customAppScopeId nor customAppScopeName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'user' || (!options.userId !== !options.userName), {
        message: "When the scope is set to 'user' specify either userId or userName, but not both",
        path: ['scope']
      })
      .refine(options => options.scope !== 'user' || (options.userId || options.userName), {
        message: "When the scope is set to 'user' specify either userId or userName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'user' || (!options.userId && !options.userName) || options.userName || (options.userId && validation.isValidGuid(options.userId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['userId']
      })
      .refine(options => options.scope !== 'user' || (!options.userId && !options.userName) || options.userId || (options.userName && validation.isValidUserPrincipalName(options.userName)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['userName']
      })
      .refine(options => options.scope !== 'group' || Object.values([options.userId, options.userName, options.administrativeUnitId, options.administrativeUnitName, options.customAppScopeId, options.customAppScopeName]).filter(v => typeof v !== 'undefined').length === 0, {
        message: "When the scope is set to 'group' then do not specify userId, userName, administrativeUnitId, administrativeUnitName, customAppScopeId nor customAppScopeName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'group' || (!options.groupId !== !options.groupName), {
        message: "When the scope is set to 'group' specify either groupId or groupName, but not both",
        path: ['scope']
      })
      .refine(options => options.scope !== 'group' || (options.groupId || options.groupName), {
        message: "When the scope is set to 'group' specify either groupId or groupName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'group' || (!options.groupId && !options.groupName) || options.groupName || (options.groupId && validation.isValidGuid(options.groupId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['groupId']
      })
      .refine(options => options.scope !== 'administrativeUnit' || Object.values([options.userId, options.userName, options.groupId, options.groupName, options.customAppScopeId, options.customAppScopeName]).filter(v => typeof v !== 'undefined').length === 0, {
        message: "When the scope is set to 'administrativeUnit' then do not specify userId, userName, groupId, groupName, customAppScopeId nor customAppScopeName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'administrativeUnit' || (!options.administrativeUnitId !== !options.administrativeUnitName), {
        message: "When the scope is set to 'administrativeUnit' specify either administrativeUnitId or administrativeUnitName, but not both",
        path: ['scope']
      })
      .refine(options => options.scope !== 'administrativeUnit' || (options.administrativeUnitId || options.administrativeUnitName), {
        message: "When the scope is set to 'administrativeUnit' specify either administrativeUnitId or administrativeUnitName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'administrativeUnit' || (!options.administrativeUnitId && !options.administrativeUnitName) || options.administrativeUnitName || (options.administrativeUnitId && validation.isValidGuid(options.administrativeUnitId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['administrativeUnitId']
      })
      .refine(options => options.scope !== 'custom' || Object.values([options.userId, options.userName, options.groupId, options.groupName, options.administrativeUnitId, options.administrativeUnitName]).filter(v => typeof v !== 'undefined').length === 0, {
        message: "When the scope is set to 'custom' then do not specify userId, userName, groupId, groupName, administrativeUnitId nor administrativeUnitName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'custom' || (!options.customAppScopeId !== !options.customAppScopeName), {
        message: "When the scope is set to 'custom' specify either customAppScopeId or customAppScopeName, but not both",
        path: ['scope']
      })
      .refine(options => options.scope !== 'custom' || (options.customAppScopeId || options.customAppScopeName), {
        message: "When the scope is set to 'custom' specify either customAppScopeId or customAppScopeName",
        path: ['scope']
      })
      .refine(options => options.scope !== 'custom' || (!options.customAppScopeId && !options.customAppScopeName) || options.customAppScopeName || (options.customAppScopeId && validation.isValidGuid(options.customAppScopeId)), {
        error: e => `The '${e.input}' must be a valid GUID`,
        path: ['customAppScopeId']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);

      const data = {
        roleDefinitionId: roleDefinitionId,
        principalId: await this.getPrincipalId(args.options, logger),
        directoryScopeId: await this.getDirectoryScopeId(args.options),
        appScopeId: await this.getAppScopeId(args.options, logger)
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

  private async getPrincipalId(options: Options, logger: Logger): Promise<string> {
    if (options.principalId) {
      return `/ServicePrincipals/${options.principalId}`;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving service principal by its name '${options.principalName}'`);
    }

    const principal = await entraServicePrincipal.getServicePrincipalByAppName(options.principalName!, 'id');
    return `/ServicePrincipals/${principal.id}`;
  }

  private async getDirectoryScopeId(options: Options): Promise<string | null> {
    if (options.scope === 'custom') {
      return null;
    }
    let prefix = '/';
    let resourceId = '';

    switch (options.scope) {
      case 'tenant':
        break;
      case 'user':
        prefix = '/users/';
        if (options.userId) {
          resourceId = options.userId;
        }
        else if (options.userName) {
          resourceId = await entraUser.getUserIdByUpn(options.userName);
        }
        break;
      case 'group':
        prefix = '/groups/';
        if (options.groupId) {
          resourceId = options.groupId;
        }
        else if (options.groupName) {
          resourceId = await entraGroup.getGroupIdByDisplayName(options.groupName);
        }
        break;
      case 'administrativeUnit':
        prefix = '/administrativeUnits/';
        if (options.administrativeUnitId) {
          resourceId = options.administrativeUnitId;
        }
        else if (options.administrativeUnitName) {
          const administrativeUnit = await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(options.administrativeUnitName);
          resourceId = administrativeUnit.id!;
        }
        break;
    }

    return `${prefix}${resourceId}`;
  }

  private async getAppScopeId(options: Options, logger: Logger): Promise<string | null> {
    if (options.scope !== 'custom') {
      return null;
    }

    if (options.customAppScopeId) {
      return options.customAppScopeId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving custom application scope by its name '${options.customAppScopeName}'`);
    }

    const applicationScopeId = (await customAppScope.getCustomAppScopeByDisplayName(options.customAppScopeName!, 'id')).id;
    return applicationScopeId!;
  }
}

export default new ExoAppRoleAssignmentAddCommand();