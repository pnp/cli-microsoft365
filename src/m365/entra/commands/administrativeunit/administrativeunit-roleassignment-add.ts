import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { roleAssignment } from '../../../../utils/roleAssignment.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  administrativeUnitId?: string;
  administrativeUnitName?: string;
  roleDefinitionId?: string;
  roleDefinitionName?: string;
  userId?: string;
  userName?: string;
}

class EntraAdministrativeUnitRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ADMINISTRATIVEUNIT_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Assigns a Microsoft Entra role with administrative unit scope to a user';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
        administrativeUnitName: typeof args.options.administrativeUnitName !== 'undefined',
        roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
        roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --administrativeUnitId [administrativeUnitId]'
      },
      {
        option: '-n, --administrativeUnitName [administrativeUnitName]'
      },
      {
        option: '--roleDefinitionId [roleDefinitionId]'
      },
      {
        option: '--roleDefinitionName [roleDefinitionName]'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.roleDefinitionId && !validation.isValidGuid(args.options.roleDefinitionId)) {
          return `${args.options.roleDefinitionId} is not a valid GUID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['administrativeUnitId', 'administrativeUnitName'] });
    this.optionSets.push({ options: ['roleDefinitionId', 'roleDefinitionName'] });
    this.optionSets.push({ options: ['userId', 'userName'] });
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