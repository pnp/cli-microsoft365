import { UnifiedRoleEligibilityScheduleInstance } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { odata } from '../../../../utils/odata.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  includePrincipalDetails?: boolean;
  withPrincipalDetails?: boolean;
}

interface UnifiedRoleEligibilityScheduleInstanceEx extends UnifiedRoleEligibilityScheduleInstance {
  roleDefinitionName?: string
}

class EntraPimRoleAssignmentEligibilityListCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_ELIGIBILITY_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of eligible roles a user or group can be assigned to';
  }

  public defaultProperties(): string[] | undefined {
    return ['roleDefinitionId', 'roleDefinitionName', 'principalId'];
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        includePrincipalDetails: !!args.options.includePrincipalDetails,
        withPrincipalDetails: !!args.options.withPrincipalDetails
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--includePrincipalDetails'
      },
      {
        option: '--withPrincipalDetails'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `'${args.options.userId} is not a valid GUID for option 'userId'.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `'${args.options.userName} is not a valid user principal name for option 'userName'.`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName', 'groupId', 'groupName'],
      runsWhen: (args) => args.options.userId || args.options.userName || args.options.groupName || args.options.groupId
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.includePrincipalDetails) {
      await this.warn(logger, `Parameter 'includePrincipalDetails' is deprecated. Please use 'withPrincipalDetails' instead`);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving list of eligible roles for ${args.options.userId || args.options.userName || args.options.groupId || args.options.groupName || 'all users'}...`);
    }
    const queryParameters: string[] = [];
    const expands: string[] = [];

    try {
      const principalId = await this.getPrincipalId(logger, args.options);

      if (principalId) {
        queryParameters.push(`$filter=principalId eq '${principalId}'`);
      }

      expands.push('roleDefinition($select=displayName)');

      if (args.options.includePrincipalDetails || args.options.withPrincipalDetails) {
        expands.push('principal');
      }

      queryParameters.push(`$expand=${expands.join(',')}`);

      const url = `${this.resource}/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?${queryParameters.join('&')}`;

      const results = await odata.getAllItems<UnifiedRoleEligibilityScheduleInstanceEx>(url);

      results.forEach(c => {
        const roleDefinition = c['roleDefinition'];

        if (roleDefinition) {
          c.roleDefinitionName = roleDefinition.displayName!;
        }

        delete c['roleDefinition'];
      });

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPrincipalId(logger: Logger, options: Options): Promise<string | undefined> {
    let principalId = options.userId;

    if (options.userName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
      }

      principalId = await entraUser.getUserIdByUpn(options.userName);
    }
    else if (options.groupId) {
      principalId = options.groupId;
    }
    else if (options.groupName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
      }

      principalId = await entraGroup.getGroupIdByDisplayName(options.groupName);
    }

    return principalId;
  }
}

export default new EntraPimRoleAssignmentEligibilityListCommand();