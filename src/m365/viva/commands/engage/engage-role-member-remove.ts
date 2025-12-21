import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { entraUser } from '../../../../utils/entraUser.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';

const options = globalOptionsZod
  .extend({
    roleId: z.string().uuid().optional(),
    roleName: z.string().optional(),
    userId: z.string().uuid().optional(),
    userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), upn => ({
      message: `'${upn}' is not a valid UPN.`
    })).optional(),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageRoleMemberRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_ROLE_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes a user from a Viva Engage role';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.roleId, options.roleName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either roleId, or roleName, but not both.'
      })
      .refine(options => [options.userId, options.userName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either userId, or userName, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeMember = async (): Promise<void> => {
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
          await logger.logToStderr(`Removing user ${userId} from a Viva Engage role ${roleId}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/employeeExperience/roles/${roleId}/members/${userId}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeMember();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user '${args.options.userId || args.options.userName}' from role '${args.options.roleId || args.options.roleName}'?` });

      if (result) {
        await removeMember();
      }
    }

  }
}

export default new VivaEngageRoleMemberRemoveCommand();