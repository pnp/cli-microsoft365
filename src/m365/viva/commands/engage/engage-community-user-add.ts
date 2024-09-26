import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';

const options = globalOptionsZod
  .extend({
    communityId: z.string().optional(),
    communityDisplayName: zod.alias('n', z.string().optional()),
    entraGroupId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    ids: z.string()
      .refine(ids => validation.isValidGuidArray(ids) === true, invalidIds => ({
        message: `The following GUIDs are invalid: ${invalidIds}.`
      })).optional(),
    userNames: z.string()
      .refine(userNames => validation.isValidUserPrincipalNameArray(userNames) === true, invalidUserNames => ({
        message: `The following user principal names are invalid: ${invalidUserNames}.`
      })).optional(),
    role: zod.alias('r', z.enum(['Admin', 'Member']))
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageCommunityUserAddCommand extends GraphCommand {

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_USER_ADD;
  }

  public get description(): string {
    return 'Adds a user to a specific Microsoft 365 Viva Engage community';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.communityId, options.communityDisplayName, options.entraGroupId].filter(x => x !== undefined).length === 1, {
        message: 'Specify either communityId, communityDisplayName, or entraGroupId, but not multiple.'
      })
      .refine(options => options.communityId || options.communityDisplayName || options.entraGroupId, {
        message: 'Specify at least one of communityId, communityDisplayName, or entraGroupId.'
      })
      .refine(options => options.ids || options.userNames, {
        message: 'Specify either of ids or userNames.'
      })
      .refine(options => typeof options.userNames !== undefined && typeof options.ids !== undefined, {
        message: 'Specify either ids or userNames, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Adding users to community...');
      }

      let entraGroupId = args.options.entraGroupId;

      if (args.options.communityDisplayName) {
        entraGroupId = await vivaEngage.getEntraGroupIdByCommunityDisplayName(args.options.communityDisplayName);
      }
      else if (args.options.communityId) {
        entraGroupId = await vivaEngage.getEntraGroupIdByCommunityId(args.options.communityId);
      }

      const userIds = args.options.ids ? formatting.splitAndTrim(args.options.ids) : await entraUser.getUserIdsByUpns(formatting.splitAndTrim(args.options.userNames!));
      const role = args.options.role === 'Member' ? 'members' : 'owners';

      for (let i = 0; i < userIds.length; i += 400) {
        const userIdsBatch = userIds.slice(i, i + 400);
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

        // only 20 requests per one batch are allowed
        for (let j = 0; j < userIdsBatch.length; j += 20) {
          // only 20 users can be added in one request
          const userIdsChunk = userIdsBatch.slice(j, j + 20);
          requestOptions.data.requests.push({
            id: j + 1,
            method: 'PATCH',
            url: `/groups/${entraGroupId}`,
            headers: {
              'content-type': 'application/json;odata.metadata=none'
            },
            body: {
              [`${role}@odata.bind`]: userIdsChunk.map((u: string) => `${this.resource}/v1.0/directoryObjects/${u}`)
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
}

export default new VivaEngageCommunityUserAddCommand();