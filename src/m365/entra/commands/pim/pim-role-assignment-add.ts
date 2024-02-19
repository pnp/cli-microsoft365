import { UnifiedRoleAssignmentScheduleRequest } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
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
  directoryScopeId?: string;
  justification?: string;
  startDateTime?: string;
  endDateTime?: string;
  duration?: string;
  ticketNumber?: string;
  ticketSystem?: string;
}

class EntraPimRoleAssignmentAddCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Request activation of an Entra ID role assignment for a user or group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.PIM_ROLE_ASSIGNMENT_ADD];
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
        roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined',
        roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        directoryScopeId: typeof args.options.directoryScopeId !== 'undefined',
        justification: typeof args.options.justification !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined',
        endDateTime: typeof args.options.endDateTime !== 'undefined',
        duration: typeof args.options.duration !== 'undefined',
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
        option: "--directoryScopeId [directoryScopeId]"
      },
      {
        option: "-j, --justification [justification]"
      },
      {
        option: "-s, --startDateTime [startDateTime]"
      },
      {
        option: "-e, --endDateTime [endDateTime]"
      },
      {
        option: "-d, --duration [duration]"
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

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return `${args.options.startDateTime} is not a valid ISO 8601 date time string`;
        }

        if (args.options.endDateTime && !validation.isValidISODateTime(args.options.endDateTime)) {
          return `${args.options.endDateTime} is not a valid ISO 8601 date time string`;
        }

        if (args.options.duration && !validation.isValidISODuration(args.options.duration)) {
          return `${args.options.duration} is not a valid ISO 8601 duration`;
        }

        if (args.options.endDateTime && args.options.duration) {
          return `Specify either 'endDateTime' or 'duration' but not both`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['roleDefinitionName', 'roleDefinitionId'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let roleDefinitionId = args.options.roleDefinitionId;
    let principalId = args.options.userId;
    let action = 'adminAssign';

    try {
      if (args.options.roleDefinitionName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving role definition by its name '${args.options.roleDefinitionName}'`);
        }

        roleDefinitionId = (await roleDefinition.getRoleDefinitionByDisplayName(args.options.roleDefinitionName)).id;
      }

      if (args.options.userName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving user by its name '${args.options.userName}'`);
        }

        principalId = await entraUser.getUserIdByUpn(args.options.userName);
      }
      else if (args.options.groupId) {
        principalId = args.options.groupId;
      }
      else if (args.options.groupName) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving group by its name '${args.options.groupName}'`);
        }

        principalId = await entraGroup.getGroupIdByDisplayName(args.options.groupName);
      }
      else if (!args.options.userId) {
        if (this.verbose) {
          await logger.logToStderr(`Retrieving id of the current user`);
        }

        const token = auth.service.accessTokens[auth.defaultResource].accessToken;
        const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);
        if (isAppOnlyAccessToken) {
          throw 'When running with application permissions either userId, userName, groupId or groupName is required';
        }

        principalId = accessToken.getUserIdFromAccessToken(token);
        action = 'selfActivate';
      }

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleRequests`,
        headers: {
          'accept': 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          principalId: principalId,
          roleDefinitionId: roleDefinitionId,
          directoryScopeId: args.options.directoryScopeId ?? '/',
          action: action,
          justification: args.options.justification,
          scheduleInfo: {
            startDateTime: args.options.startDateTime,
            expiration: {
              duration: args.options.duration,
              endDateTime: args.options.endDateTime,
              type: this.getExpirationType(args.options)
            }
          },
          ticketInfo: {
            ticketNumber: args.options.ticketNumber,
            ticketSystem: args.options.ticketSystem
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

  private getExpirationType(options: Options): string {
    if (options.duration) {
      return 'afterDuration';
    }

    if (options.endDateTime) {
      return 'afterDateTime';
    }

    return 'noExpiration';
  }
}

export default new EntraPimRoleAssignmentAddCommand();