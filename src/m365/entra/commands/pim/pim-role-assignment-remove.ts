import { UnifiedRoleAssignmentScheduleRequest } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { roleDefinition } from '../../../../utils/roleDefinition.js';
import { validation } from '../../../../utils/validation.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  roleDefinitionName?: string;
  roleDefinitionId?: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  administrativeUnitId?: string;
  applicationId?: string;
  justification?: string,
  ticketNumber?: string;
  ticketSystem?: string;
}

class EntraPimRoleAssignmentRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Request deactivation of an Entra role assignment for a user or group';
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
        roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined',
        roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        administrativeUnitId: typeof args.options.administrativeUnitId !== 'undefined',
        applicationId: typeof args.options.applicationId !== 'undefined',
        justification: typeof args.options.justification !== 'undefined',
        ticketNumber: typeof args.options.ticketNumber !== 'undefined',
        ticketSystem: typeof args.options.ticketSystem !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --roleDefinitionName [roleDefinitionName]'
      },
      {
        option: '-i, --roleDefinitionId [roleDefinitionId]'
      },
      {
        option: "--userId [userId]"
      },
      {
        option: "--userName [userName]"
      },
      {
        option: "--groupId [groupId]"
      },
      {
        option: "--groupName [groupName]"
      },
      {
        option: "--administrativeUnitId [administrativeUnitId]"
      },
      {
        option: "--applicationId [applicationId]"
      },
      {
        option: "-j, --justification [justification]"
      },
      {
        option: "--ticketNumber [ticketNumber]"
      },
      {
        option: "--ticketSystem [ticketSystem]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.roleDefinitionId && !validation.isValidGuid(args.options.roleDefinitionId)) {
          return `${args.options.roleDefinitionId} is not a valid GUID`;
        }

        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.administrativeUnitId && !validation.isValidGuid(args.options.administrativeUnitId)) {
          return `${args.options.administrativeUnitId} is not a valid GUID`;
        }

        if (args.options.applicationId && !validation.isValidGuid(args.options.applicationId)) {
          return `${args.options.applicationId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['roleDefinitionName', 'roleDefinitionId'] });
    this.optionSets.push({
      options: ['userId', 'userName', 'groupId', 'groupName'],
      runsWhen: (args) => {
        return args.options.userId !== undefined || args.options.userName !== undefined || args.options.groupId !== undefined || args.options.groupName !== undefined;
      }
    });
    this.optionSets.push({
      options: ['administrativeUnitId', 'applicationId'],
      runsWhen: (args) => {
        return args.options.administrativeUnitId !== undefined || args.options.applicationId !== undefined;
      }
    });
  }

  #initTypes(): void {
    this.types.string.push('userId', 'userName', 'groupId', 'groupName', 'administrativeUnitId', 'applicationId', 'roleDefinitionName', 'roleDefinitionId', 'justification', 'ticketNumber', 'ticketSystem');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const { userId, userName, groupId, groupName, ticketNumber, ticketSystem } = args.options;
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      if (isAppOnlyAccessToken && !userId && !userName && !groupId && !groupName) {
        throw 'When running with application permissions either userId, userName, groupId or groupName is required';
      }

      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      const principalId = await this.getPrincipalId(args.options, logger);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          principalId: principalId,
          roleDefinitionId: roleDefinitionId,
          directoryScopeId: this.getDirectoryScope(args.options),
          action: !userId && !userName && !groupId && !groupName ? 'selfDeactivate' : 'adminRemove',
          justification: args.options.justification,
          ticketInfo: {
            ticketNumber: ticketNumber,
            ticketSystem: ticketSystem
          }
        }
      };

      const response = await request.post<UnifiedRoleAssignmentScheduleRequest>(requestOptions);

      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<string> {
    if (options.roleDefinitionId) {
      return options.roleDefinitionId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving role definition by its name '${options.roleDefinitionName}'`);
    }

    const role = await roleDefinition.getRoleDefinitionByDisplayName(options.roleDefinitionName!);
    return role.id!;
  }

  private async getPrincipalId(options: Options, logger: Logger): Promise<string> {
    if (options.userId || options.groupId) {
      return options.userId! || options.groupId!;
    }

    if (options.userName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
      }

      return await entraUser.getUserIdByUpn(options.userName);
    }
    else if (options.groupName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
      }

      return await entraGroup.getGroupIdByDisplayName(options.groupName);
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving id of the current user`);
    }

    const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
    return accessToken.getUserIdFromAccessToken(token);
  }

  private getDirectoryScope(options: Options): string {
    if (options.administrativeUnitId) {
      return `/administrativeUnits/${options.administrativeUnitId}`;
    }

    if (options.applicationId) {
      return `/${options.applicationId}`;
    }

    return '/';
  }
}

export default new EntraPimRoleAssignmentRemoveCommand();