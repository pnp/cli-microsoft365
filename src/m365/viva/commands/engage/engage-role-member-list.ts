import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';

const options = globalOptionsZod
  .extend({
    roleId: zod.alias('i', z.string().refine(name => validation.isValidGuid(name), name => ({
      message: `'${name}' is not a valid GUID.`
    })).optional()),
    roleName: zod.alias('n', z.string().optional())
  })
  .strict();
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class VivaEngageRoleMemberListCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_ROLE_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists all users assigned to a Viva Engage role';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.roleId, options.roleName].filter(x => x !== undefined).length === 1, {
        message: 'Specify either roleId, or roleName, but not both.'
      });
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'userId'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let roleId = args.options.roleId;

    try {
      if (args.options.roleName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving Viva Engage role ID for role name '${args.options.roleName}'...`);
        }
        roleId = await vivaEngage.getRoleIdByName(args.options.roleName);
      }

      if (this.verbose) {
        await logger.logToStderr(`Getting all users assigned to a Viva Engage role ${roleId}...`);
      }

      const results = await odata.getAllItems<any>(`${this.resource}/beta/employeeExperience/roles/${roleId}/members`);
      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new VivaEngageRoleMemberListCommand();