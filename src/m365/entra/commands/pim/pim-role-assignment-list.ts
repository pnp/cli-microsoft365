import { UnifiedRoleAssignmentScheduleInstance } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
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
  startDateTime?: string;
  includePrincipalDetails?: boolean;
}

class EntraPimRoleAssignmentListCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of Entra role assignments for a user or group';
  }


  public alias(): string[] | undefined {
    return [aadCommands.PIM_ROLE_ASSIGNMENT_LIST];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined',
        includePrincipalDetails: !!args.options.includePrincipalDetails
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: "-s, --startDateTime [startDateTime]"
      },
      {
        option: "--includePrincipalDetails [includePrincipalDetails]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return `${args.options.startDateTime} is not a valid ISO 8601 date time string`;
        }

        if (args.options.userId && args.options.userName || args.options.userId && args.options.groupId ||
          args.options.userId && args.options.groupName || args.options.userName && args.options.groupId ||
          args.options.userName && args.options.groupName || args.options.groupId && args.options.groupName) {
          return `Specify either userId, userName, groupId or groupName.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let principalId = args.options.userId;

    try {
      const queryParameters: string[] = [];

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
      
      const filters: string[] = [];
      if (principalId) {
        filters.push(`principalId eq '${principalId}'`);
      }

      if (args.options.startDateTime) {
        filters.push(`startDateTime ge ${args.options.startDateTime}`);
      }

      if (filters.length > 0) {
        queryParameters.push(`$filter=${filters.join(' and ')}`);
      }

      if (args.options.includePrincipalDetails) {
        queryParameters.push('$expand=principal');
      }

      const queryString = queryParameters.length > 0
        ? `?${queryParameters.join('&')}`
        : '';

      const url = `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleInstances${queryString}`;

      const results = await odata.getAllItems<UnifiedRoleAssignmentScheduleInstance>(url);

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraPimRoleAssignmentListCommand();