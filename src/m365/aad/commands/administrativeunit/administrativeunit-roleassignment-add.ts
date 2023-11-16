import GlobalOptions from "../../../../GlobalOptions.js";
import { Logger } from "../../../../cli/Logger.js";
import { aadAdministrativeUnit } from "../../../../utils/aadAdministrativeUnit.js";
import { aadUser } from "../../../../utils/aadUser.js";
import { roleAssignment } from "../../../../utils/roleAssignment.js";
import { roleDefinition } from "../../../../utils/roleDefinition.js";
import { validation } from "../../../../utils/validation.js";
import GraphCommand from "../../../base/GraphCommand.js";
import commands from "../../commands.js";

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

class AadAdministrativeUnitRoleAssignmentAddCommand extends GraphCommand {
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
    this.#initTypes();
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
        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId as string)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.roleDefinitionId && !validation.isValidGuid(args.options.roleDefinitionId as string)) {
          return `${args.options.roleDefinitionId} is not a valid GUID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId as string)) {
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

  #initTypes(): void {
    this.types.string.push('administrativeUnitName', 'roleName', 'userName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let administrativeUnitId = args.options.administrativeUnitId;
      let roleDefinitionId = args.options.roleDefinitionId;
      let userId = args.options.userId;

      if (args.options.administrativeUnitName) {
        administrativeUnitId = (await aadAdministrativeUnit.getAdministrativeUnitByDisplayName(args.options.administrativeUnitName)).id;
      }

      if (args.options.roleDefinitionName) {
        roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.roleDefinitionName)).id;
      }

      if (args.options.userName) {
        userId = await aadUser.getUserIdByUpn(args.options.userName);
      }

      const unifiedRoleAssignment = await roleAssignment.createRoleAssignmentWithAdministrativeUnitScope(roleDefinitionId!, userId!, administrativeUnitId!);

      await logger.log(unifiedRoleAssignment);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new AadAdministrativeUnitRoleAssignmentAddCommand();