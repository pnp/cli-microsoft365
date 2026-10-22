import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
      error: e => `The following GUIDs are invalid for the option 'userIds': ${validation.isValidGuidArray(e.input as string)}.`
    }).optional(),
  userNames: z.string()
    .refine(names => validation.isValidUserPrincipalNameArray(names) === true, {
      error: e => `The following user principal names are invalid for the option 'userNames': ${validation.isValidUserPrincipalNameArray(e.input as string)}.`
    }).optional(),
  subgroupIds: z.string()
    .refine(ids => validation.isValidGuidArray(ids) === true, {
      error: e => `The following GUIDs are invalid for the option 'subgroupIds': ${validation.isValidGuidArray(e.input as string)}.`
    }).optional(),
  subgroupNames: z.string().optional(),
  role: zod.coercedEnum(RoleEnum).alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraGroupMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds members to a Microsoft Entra group';
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
      .refine(options => !(options.subgroupIds || options.subgroupNames) || options.role !== 'Owner', {
        error: `Subgroups cannot be set as owners.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Adding member(s) ${args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames} to group ${ args.options.groupId || args.options.groupName }...`);
      }

      const groupId = await this.getGroupId(logger, args.options);
      const objectIds = await this.getObjectIds(logger, args.options);

      for (let i = 0; i < objectIds.length; i += 400) {
        const objectIdsBatch = objectIds.slice(i, i + 400);
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/$batch`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            requests: []
          }
        };

        for (let j = 0; j < objectIdsBatch.length; j += 20) {
          const objectIdsChunk = objectIdsBatch.slice(j, j + 20);
          requestOptions.data.requests.push({
            id: j + 1,
            method: 'PATCH',
            url: `/groups/${groupId}`,
            headers: {
              'content-type': 'application/json;odata.metadata=none'
            },
            body: {
              [`${args.options.role === 'Member' ? 'members' : 'owners'}@odata.bind`]: objectIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
            }
          });
        }

        const res = await request.post<{ responses: { status: number; body: any }[] }>(requestOptions);
        for (const response of res.responses) {
          if (response.status !== 204) {
            throw response.body;
          }
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
      await logger.logToStderr(`Retrieving ID of group ${options.groupName}...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupName!);
  }

  private async getObjectIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.userIds || options.userNames) {
      return this.getUserIds(logger, options);
    }

    return this.getGroupIds(logger, options);
  }

  private async getUserIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.userIds) {
      return options.userIds.split(',').map(i => i.trim());
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of user(s)...');
    }

    return entraUser.getUserIdsByUpns(options.userNames!.split(',').map(u => u.trim()));
  }

  private async getGroupIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.subgroupIds) {
      return options.subgroupIds.split(',').map(i => i.trim());
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of group(s)...');
    }

    return entraGroup.getGroupIdsByDisplayNames(options.subgroupNames!.split(',').map(u => u.trim()));
  }
}

export default new EntraGroupMemberAddCommand();