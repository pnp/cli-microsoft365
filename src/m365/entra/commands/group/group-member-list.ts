import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';
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
  role: zod.coercedEnum(RoleEnum).optional().alias('r'),
  properties: z.string().optional().alias('p'),
  filter: z.string().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ExtendedUser extends User {
  roles: string[];
}

class EntraGroupMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists members of a specific Entra group';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'userPrincipalName', 'roles'];
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args.options, logger);

      const users: ExtendedUser[] = [];

      if (!args.options.role || args.options.role === 'Owner') {
        const owners = await this.getUsers(args.options, 'Owners', groupId, logger);
        owners.forEach(owner => users.push({ ...owner, roles: ['Owner'] }));
      }

      if (!args.options.role || args.options.role === 'Member') {
        const members = await this.getUsers(args.options, 'Members', groupId, logger);

        members.forEach((member: ExtendedUser) => {
          const user = users.find((u: ExtendedUser) => u.id === member.id);

          if (user !== undefined) {
            user.roles.push('Member');
          }
          else {
            users.push({ ...member, roles: ['Member'] });
          }
        });
      }

      await logger.log(users);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(options: Options, logger: Logger): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving Group Id...');
    }

    return await entraGroup.getGroupIdByDisplayName(options.groupName!);
  }

  private async getUsers(options: Options, role: string, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
      'id,displayName,userPrincipalName,givenName,surname';
    const allSelectProperties: string[] = selectProperties.split(',');
    const propertiesWithSlash: string[] = allSelectProperties.filter(item => item.includes('/'));

    let fieldExpand: string = '';
    propertiesWithSlash.forEach(p => {
      if (fieldExpand.length > 0) {
        fieldExpand += ',';
      }

      fieldExpand += `${p.split('/')[0]}($select=${p.split('/')[1]})`;
    });

    const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
    const selectParam = allSelectProperties.filter(item => !item.includes('/'));
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${role}?$select=${selectParam}${expandParam}`;

    let users: ExtendedUser[];

    if (filter) {
      // While using the filter, we need to specify the ConsistencyLevel header.
      // Can be refactored when the header is no longer necessary.
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}&$filter=${encodeURIComponent(filter)}&$count=true`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          ConsistencyLevel: 'eventual'
        },
        responseType: 'json'
      };

      users = await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      users = await odata.getAllItems<ExtendedUser>(endpoint);
    }

    return users;
  }
}

export default new EntraGroupMemberListCommand();