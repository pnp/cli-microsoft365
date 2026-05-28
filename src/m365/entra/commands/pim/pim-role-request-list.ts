import { UnifiedRoleAssignmentScheduleRequest } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';

const allowedStatuses = ['Canceled', 'Denied', 'Failed', 'Granted', 'PendingAdminDecision', 'PendingApproval', 'PendingProvisioning', 'PendingScheduleCreation', 'Provisioned', 'Revoked', 'ScheduleCreated'] as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional(),
  userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), {
    error: e => `'${e.input}' is not a valid user principal name for option 'userName'.`
  }).optional(),
  groupId: z.uuid().optional(),
  groupName: z.string().optional(),
  createdDateTime: z.string().refine(date => validation.isValidISODateTime(date), {
    error: e => `'${e.input}' is not a valid ISO 8601 date time string for option 'createdDateTime'.`
  }).optional().alias('c'),
  status: z.enum(allowedStatuses).optional().alias('s'),
  withPrincipalDetails: z.boolean().default(false)
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface UnifiedRoleAssignmentScheduleRequestEx extends UnifiedRoleAssignmentScheduleRequest {
  roleDefinitionName?: string
}

class EntraPimRoleRequestListCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_REQUEST_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of PIM requests for roles';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(options => {
        const specified = [options.userId, options.userName, options.groupId, options.groupName].filter(o => o !== undefined).length;
        return specified <= 1;
      }, {
        message: 'Specify only one of the following options: userId, userName, groupId, groupName',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName']
        }
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'roleDefinitionName', 'principalId'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of PIM roles requests for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || 'all users'}...`);
    }
    const queryParameters: string[] = [];
    const filters: string[] = [];
    const expands: string[] = [];

    try {
      const principalId = await this.getPrincipalId(logger, args.options);

      if (principalId) {
        filters.push(`principalId eq '${principalId}'`);
      }

      if (args.options.createdDateTime) {
        filters.push(`createdDateTime ge ${args.options.createdDateTime}`);
      }

      if (args.options.status) {
        filters.push(`status eq '${args.options.status}'`);
      }

      if (filters.length > 0) {
        queryParameters.push(`$filter=${filters.join(' and ')}`);
      }

      expands.push('roleDefinition($select=displayName)');

      if (args.options.withPrincipalDetails) {
        expands.push('principal');
      }

      queryParameters.push(`$expand=${expands.join(',')}`);

      const queryString = `?${queryParameters.join('&')}`;

      const url = `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests${queryString}`;

      const results = await odata.getAllItems<UnifiedRoleAssignmentScheduleRequestEx>(url);

      results.forEach(c => {
        const roleDefinition = c['roleDefinition'];

        if (roleDefinition) {
          c.roleDefinitionName = roleDefinition.displayName!;
        }

        delete c['roleDefinition'];
      });

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPrincipalId(logger: Logger, options: Options): Promise<string | undefined> {
    let principalId = options.userId;

    if (options.userName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
      }

      principalId = await entraUser.getUserIdByUpn(options.userName);
    }
    else if (options.groupId) {
      principalId = options.groupId;
    }
    else if (options.groupName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
      }

      principalId = await entraGroup.getGroupIdByDisplayName(options.groupName);
    }

    return principalId;
  }
}

export default new EntraPimRoleRequestListCommand();