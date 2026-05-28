import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { GraphBatchRequest, GraphBatchRequestResponse } from '../../../../utils/types.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const RoleEnum = {
  Owner: 'Owner',
  Member: 'Member'
} as const;

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  groupId: z.uuid().optional().alias('i'),
  groupName: z.string().optional().alias('n'),
  userIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `Invalid GUIDs found for option 'ids': ${e.input}.`
    }).optional(),
  userNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `Invalid UPNs found for option 'userNames': ${e.input}.`
    }).optional(),
  subgroupIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `Invalid GUIDs found for option 'subgroupIds': ${e.input}.`
    }).optional(),
  subgroupNames: z.string().optional(),
  role: zod.coercedEnum(RoleEnum).optional().alias('r'),
  suppressNotFound: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes members from a Microsoft Entra group';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.groupId, options.groupName].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: groupId or groupName.',
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName']
        }
      })
      .refine(options => [options.userIds, options.userNames, options.subgroupIds, options.subgroupNames].filter(o => o !== undefined).length === 1, {
        error: 'Use one of the following options: userIds, userNames, subgroupIds, or subgroupNames.',
        params: {
          customCode: 'optionSet',
          options: ['userIds', 'userNames', 'subgroupIds', 'subgroupNames']
        }
      })
      .refine(options => !(options.subgroupIds !== undefined || options.subgroupNames !== undefined) || options.role?.toLowerCase() === 'member', {
        error: `When removing subgroups, the 'role' option must be set to 'Member'.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const removeUsers = async (): Promise<void> => {
        if (this.verbose) {
          await logger.logToStderr(`Removing user(s) ${args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames} from group ${args.options.groupId || args.options.groupName}...`);
        }

        const groupId = await this.getGroupId(logger, args.options);
        const userIds = await this.getPrincipalIds(logger, args.options);

        const endpoints = [];
        if (!args.options.role || args.options.role === 'Owner') {
          endpoints.push(...userIds.map(id => `/groups/${groupId}/owners/${id}/$ref`));
        }
        if (!args.options.role || args.options.role === 'Member') {
          endpoints.push(...userIds.map(id => `/groups/${groupId}/members/${id}/$ref`));
        }

        for (let i = 0; i < endpoints.length; i += 20) {
          const endpointsBatch = endpoints.slice(i, i + 20);
          const requestOptions: CliRequestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
              'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
              requests: endpointsBatch.map((ep, index) => ({
                id: index + 1,
                method: 'DELETE',
                url: ep,
                headers: {
                  'content-type': 'application/json;odata.metadata=none'
                }
              }))
            } as GraphBatchRequest
          };

          const res = await request.post<GraphBatchRequestResponse>(requestOptions);
          for (const response of res.responses) {
            // Suppress 404 errors if suppressNotFound is set
            if (response.status !== 204 && (!args.options.suppressNotFound || response.status !== 404)) {
              throw response.body;
            }
          }
        }
      };

      if (args.options.force) {
        await removeUsers();
      }
      else {
        const principals = args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames;
        const principalsList = principals!.split(',');
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${principalsList.length} principal(s) from group '${args.options.groupId || args.options.groupName}'?` });

        if (result) {
          await removeUsers();
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(logger: Logger, options: Options): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ID of group '${options.groupName}'...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupName!);
  }

  private async getPrincipalIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.userIds) {
      return options.userIds.split(',').map(i => i.trim());
    }

    if (options.subgroupIds) {
      return options.subgroupIds.split(',').map(i => i.trim());
    }

    if (options.userNames) {
      if (this.verbose) {
        await logger.logToStderr('Retrieving ID(s) of user(s)...');
      }

      return entraUser.getUserIdsByUpns(options.userNames!.split(',').map(u => u.trim()));
    }

    // Subgroup names were specified
    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of subgroup(s)...');
    }

    const subGroupIds: string[] = [];
    for (const subgroupName of options.subgroupNames!.split(',')) {
      const groupId = await entraGroup.getGroupIdByDisplayName(subgroupName.trim());
      subGroupIds.push(groupId);
    }
    return subGroupIds;
  }
}

export default new EntraGroupMemberRemoveCommand();