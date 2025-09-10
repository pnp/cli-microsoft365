import { User } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { CliRequestOptions } from '../../../../request.js';

const options = globalOptionsZod
  .extend({
    filter: zod.alias('f', z.string().optional()),
    groupId: zod.alias('i', z.string().uuid().optional()),
    groupDisplayName: zod.alias('d', z.string().optional()),
    properties: zod.alias('p', z.string().optional()),
    role: zod.alias('r', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ExtendedUser extends User {
  roles: string[];
}

class EntraM365GroupUserListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_USER_LIST;
  }

  public get description(): string {
    return "Lists users for the specified Microsoft 365 group";
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !!(options.groupId || options.groupDisplayName), {
        message: 'Specify either groupId or groupDisplayName'
      })
      .refine(options => {
        if (options.role && !['Owner', 'Member'].includes(options.role)) {
          return false;
        }
        return true;
      }, {
        message: 'Invalid role value. Allowed values Owner|Member'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const groupId = await this.getGroupId(args.options, logger);
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group '${args.options.groupId || args.options.groupDisplayName}' is not a Microsoft 365 group.`);
      }

      let users: ExtendedUser[] = [];
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

      if (args.options.role) {
        users = users.filter(i => i.roles.indexOf(args.options.role!) > -1);
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

    return await entraGroup.getGroupIdByDisplayName(options.groupDisplayName!);
  }

  private async getUsers(options: Options, role: string, groupId: string, logger: Logger): Promise<ExtendedUser[]> {
    const { properties, filter } = options;

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ${role} of the group with id ${groupId}`);
    }

    const selectProperties: string = properties ?
      `${properties.split(',').filter(f => f.toLowerCase() !== 'id').concat('id').map(p => p.trim()).join(',')}` :
      'id,displayName,userPrincipalName,givenName,surname,userType';
    const allSelectProperties: string[] = selectProperties.split(',');
    const propertiesWithSlash: string[] = allSelectProperties.filter(item => item.includes('/'));

    const fieldsToExpand: string[] = [];
    propertiesWithSlash.forEach(p => {
      const propertiesSplit: string[] = p.split('/');
      fieldsToExpand.push(`${propertiesSplit[0]}($select=${propertiesSplit[1]})`);
    });

    const fieldExpand: string = fieldsToExpand.join(',');

    const expandParam = fieldExpand.length > 0 ? `&$expand=${fieldExpand}` : '';
    const selectParam = allSelectProperties.filter(item => !item.includes('/'));
    const endpoint: string = `${this.resource}/v1.0/groups/${groupId}/${role}/microsoft.graph.user?$select=${selectParam}${expandParam}`;

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

      return await odata.getAllItems<ExtendedUser>(requestOptions);
    }
    else {
      return await odata.getAllItems<ExtendedUser>(endpoint);
    }
  }
}

export default new EntraM365GroupUserListCommand();