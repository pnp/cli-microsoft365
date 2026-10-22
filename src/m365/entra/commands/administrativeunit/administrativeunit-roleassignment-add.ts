import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { roleAssignment } from '../../../../utils/roleAssignment.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  administrativeUnitId: z.uuid().optional().alias('i'),
  administrativeUnitName: z.string().optional().alias('n'),
  roleDefinitionId: z.uuid().optional(),
  roleDefinitionName: z.string().optional(),
  userId: z.uuid().optional(),
  userName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraAdministrativeUnitRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Assigns a Microsoft Entra role with administrative unit scope to a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.administrativeUnitId, options.administrativeUnitName].filter(Boolean).length === 1, {
        error: 'Specify either administrativeUnitId or administrativeUnitName',
        params: {
          customCode: 'optionSet',
          options: ['administrativeUnitId', 'administrativeUnitName']
        }
      })
      .refine(options => [options.roleDefinitionId, options.roleDefinitionName].filter(Boolean).length === 1, {
        error: 'Specify either roleDefinitionId or roleDefinitionName',
        params: {
          customCode: 'optionSet',
          options: ['roleDefinitionId', 'roleDefinitionName']
        }
      })
      .refine(options => [options.userId, options.userName].filter(Boolean).length === 1, {
        error: 'Specify either userId or userName',
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let { administrativeUnitId, roleDefinitionId, userId } = args.options;

      if (args.options.administrativeUnitName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving administrative unit by its name '${args.options.administrativeUnitName}'`);
        }

        administrativeUnitId = (await entraAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
      }

      if (args.options.roleDefinitionName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving role definition by its name '${args.options.roleDefinitionName}'`);
        }

        roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.roleDefinitionName)).id;
      }

      if (args.options.userName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving user by UPN '${args.options.userName}'`);
        }

        userId = await entraUser.getUserIdByUpn(args.options.userName);
      }

      const unifiedRoleAssignment = await roleAssignment.createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId!, userId!, administrativeUnitId!);

      await logger.log(unifiedRoleAssignment);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraAdministrativeUnitRoleAssignmentAddCommand();