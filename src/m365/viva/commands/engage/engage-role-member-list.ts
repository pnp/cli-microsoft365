import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  roleId: z.uuid().optional().alias('i'),
  roleName: z.string().optional().alias('n')
});
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

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.roleId, options.roleName].filter(x => x !== undefined).length === 1, {
        error: 'Specify either roleId, or roleName, but not both.'
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