import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { ODataResponse } from '../../../../utils/odata.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional().alias('i'),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid user principal name.`
  }).optional().alias('n'),
  userEmail: z.string().refine(email => validation.isValidUserPrincipalName(email), {
    error: e => `'${e.input}' is not a valid user email.`
  }).optional().alias('e'),
  securityEnabledOnly: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface UserGroupMembership {
  groupId: string;
}

class EntraUserGroupmembershipListCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_GROUPMEMBERSHIP_LIST;
  }

  public get description(): string {
    return 'Retrieves all groups where the user is a member of';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName, options.userEmail].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'userId', 'userName', or 'userEmail'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'userEmail']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let userId = args.options.userId;

    try {
      if (args.options.userName) {
        userId = await entraUser.getUserIdByUpn(args.options.userName);
      }
      else if (args.options.userEmail) {
        userId = await entraUser.getUserIdByEmail(args.options.userEmail);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${userId}/getMemberGroups`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          securityEnabledOnly: !!args.options.securityEnabledOnly
        }
      };

      const groups: UserGroupMembership[] = [];

      const results = await request.post<ODataResponse<string>>(requestOptions);

      results.value.forEach(x => groups.push({ groupId: x }));

      await logger.log(groups);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraUserGroupmembershipListCommand();