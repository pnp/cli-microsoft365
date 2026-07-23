import { UnifiedRoleAssignmentScheduleRequest } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  roleDefinitionName: z.string().optional().alias('n'),
  roleDefinitionId: z.uuid().optional().alias('i'),
  userId: z.uuid().optional(),
  userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), {
    error: e => `'${e.input}' is not a valid user principal name for option 'userName'.`
  }).optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  administrativeUnitId: z.uuid().optional(),
  applicationId: z.uuid().optional(),
  justification: z.string().optional().alias('j'),
  ticketNumber: z.string().optional(),
  ticketSystem: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraPimRoleAssignmentRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Requests deactivation of an Entra role assignment for a user or group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(options => [options.roleDefinitionId, options.roleDefinitionName].filter(o => o !== undefined).length === 1, {
        message: 'Specify either roleDefinitionId or roleDefinitionName',
        params: {
          customCode: 'optionSet',
          options: ['roleDefinitionId', 'roleDefinitionName']
        }
      })
      .refine(options => {
        const specified = [options.userId, options.userName, options.groupId, options.groupName].filter(o => o !== undefined).length;
        return specified <= 1;
      }, {
        message: 'Specify only one of the following options: userId, userName, groupId, groupName',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName']
        }
      })
      .refine(options => {
        const specified = [options.administrativeUnitId, options.applicationId].filter(o => o !== undefined).length;
        return specified <= 1;
      }, {
        message: 'Specify only one of the following options: administrativeUnitId, applicationId',
        params: {
          customCode: 'optionSet',
          options: ['administrativeUnitId', 'applicationId']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { userId, userName, groupId, groupName, ticketNumber, ticketSystem } = args.options;
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      if (isAppOnlyAccessToken && !userId && !userName && !groupId && !groupName) {
        throw 'When running with application permissions either userId, userName, groupId or groupName is required';
      }

      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      const principalId = await this.getPrincipalId(args.options, logger);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          principalId: principalId,
          roleDefinitionId: roleDefinitionId,
          directoryScopeId: this.getDirectoryScope(args.options),
          action: !userId && !userName && !groupId && !groupName ? 'selfDeactivate' : 'adminRemove',
          justification: args.options.justification,
          ticketInfo: {
            ticketNumber: ticketNumber,
            ticketSystem: ticketSystem
          }
        }
      };

      const response = await request.post<UnifiedRoleAssignmentScheduleRequest>(requestOptions);

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
    if (options.userId || options.groupId) {
      return options.userId! || options.groupId!;
    }

    if (options.userName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
      }

      return await entraUser.getUserIdByUpn(options.userName);
    }
    else if (options.groupName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
      }

      return await entraGroup.getGroupIdByDisplayName(options.groupName);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving id of the current user`);
    }

    const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
    return accessToken.getUserIdFromAccessToken(token);
  }

  private getDirectoryScope(options: Options): string {
    if (options.administrativeUnitId) {
      return `/administrativeUnits/${options.administrativeUnitId}`;
    }

    if (options.applicationId) {
      return `/${options.applicationId}`;
    }

    return '/';
  }
}

export default new EntraPimRoleAssignmentRemoveCommand();