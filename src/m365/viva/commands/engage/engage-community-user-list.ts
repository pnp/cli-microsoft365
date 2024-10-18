import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { CliRequestOptions } from '../../../../request.js';
import { User } from '@microsoft/microsoft-graph-types';
import { odata } from '../../../../utils/odata.js';

const options = globalOptionsZod
  .extend({
    communityId: z.string().optional(),
    communityDisplayName: zod.alias('n', z.string().optional()),
    entraGroupId: z.string()
      .refine(name => validation.isValidGuid(name), name => ({
        message: `'${name}' is not a valid GUID.`
      })).optional(),
    role: zod.alias('r', z.enum(['Admin', 'Member']).optional())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ExtendedUser extends User {
  roles: string[];
}

class VivaEngageCommunityUserListCommand extends GraphCommand {

  public get name(): string {
    return commands.ENGAGE_COMMUNITY_USER_LIST;
  }

  public get description(): string {
    return 'Lists all users within a specified Microsoft 365 Viva Engage community';
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
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName', 'roles'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr('Getting list of users in community...');
      }

      let entraGroupId = args.options.entraGroupId;

      if (args.options.communityDisplayName) {
        const community = await vivaEngage.getCommunityByDisplayName(args.options.communityDisplayName, ['groupId']);
        entraGroupId = community.groupId;
      }

      if (args.options.communityId) {
        const community = await vivaEngage.getCommunityById(args.options.communityId, ['groupId']);
        entraGroupId = community.groupId;
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/groups/${entraGroupId}/members`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };
      const members = await odata.getAllItems<User[]>(requestOptions);

      requestOptions.url = `${this.resource}/v1.0/groups/${entraGroupId}/owners`;
      const owners = await odata.getAllItems<User[]>(requestOptions);

      const extendedMembers: ExtendedUser[] = members.map(m => {
        return {
          ...m,
          roles: ['Member']
        };
      });

      const extendedOwners: ExtendedUser[] = owners.map(o => {
        return {
          ...o,
          roles: ['Admin']
        };
      });

      let users: ExtendedUser[] = [];
      if (args.options.role) {
        if (args.options.role === 'Member') {
          users = users.concat(extendedMembers);
        }
        if (args.options.role === 'Admin') {
          users = users.concat(extendedOwners);
        }
      }
      else {
        users = extendedOwners.concat(extendedMembers);
      }

      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageCommunityUserListCommand();