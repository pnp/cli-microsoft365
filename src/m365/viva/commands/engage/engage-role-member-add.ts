import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  roleId: z.uuid().optional().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }),
  roleName: z.string().optional(),
  userId: z.uuid().optional().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }),
  userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageRoleMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_ROLE_MEMBER_ADD;
  }

  public get description(): string {
    return 'Assigns a Viva Engage role to a user';
  }

  public get schema(): z.ZodType<Options> {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.roleId, options.roleName].filter(x => x !== undefined).length === 1, {
        error: 'Specify either roleId, or roleName, but not both.'
      })
      .refine(options => [options.userId, options.userName].filter(x => x !== undefined).length === 1, {
        error: 'Specify either userId, or userName, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let roleId = args.options.roleId;
      let userId = args.options.userId;

      if (args.options.roleName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving Viva Engage role ID for role name '${args.options.roleName}'...`);
        }
        roleId = await vivaEngage.getRoleIdByName(args.options.roleName);
      }

      if (args.options.userName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving Viva Engage user ID for user name '${args.options.userName}'...`);
        }
        userId = await entraUser.getUserIdByUpn(args.options.userName);
      }

      if (this.verbose) {
        await logger.logToStderr(`Assigning user ${userId} to a Viva Engage role ${roleId}...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/employeeExperience/roles/${roleId}/members`,
        headers: {
          'content-type': 'application/json;odata=nometadata',
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          'user@odata.bind': `${this.resource}/beta/users('${userId}')`
        }
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageRoleMemberAddCommand();