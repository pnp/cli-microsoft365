import GlobalOptions from '../../../../GlobalOptions.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { GraphBatchRequest, GraphBatchRequestResponse } from '../../../../utils/types.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupName?: string;
  userIds?: string;
  userNames?: string;
  subgroupIds?: string;
  subgroupNames?: string;
  role?: string;
  suppressNotFound?: boolean;
  force?: boolean;
}

class EntraGroupMemberRemoveCommand extends GraphCommand {
  private readonly roleValues = ['Owner', 'Member'];

  public get name(): string {
    return commands.GROUP_MEMBER_REMOVE;
  }

  public get description(): string {
    return 'Removes members from a Microsoft Entra group';
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
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        userIds: typeof args.options.userIds !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined',
        subgroupIds: typeof args.options.subgroupIds !== 'undefined',
        subgroupNames: typeof args.options.subgroupNames !== 'undefined',
        role: typeof args.options.role !== 'undefined',
        suppressNotFound: !!args.options.suppressNotFound,
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId [groupId]'
      },
      {
        option: '-n, --groupName [groupName]'
      },
      {
        option: '--userIds [userIds]'
      },
      {
        option: '--userNames [userNames]'
      },
      {
        option: '--subgroupIds [subgroupIds]'
      },
      {
        option: '--subgroupNames [subgroupNames]'
      },
      {
        option: '-r, --role [role]',
        autocomplete: this.roleValues
      },
      {
        option: '--suppressNotFound'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId !== undefined && !validation.isValidGuid(args.options.groupId)) {
          return `'${args.options.groupId}' is not a valid GUID for option 'groupId'.`;
        }

        if (args.options.userIds !== undefined) {
          const invalidGuids = validation.isValidGuidArray(args.options.userIds);
          if (invalidGuids !== true) {
            return `Invalid GUIDs found for option 'ids': ${invalidGuids}.`;
          }
        }

        if (args.options.userNames !== undefined) {
          const invalidUpns = validation.isValidUserPrincipalNameArray(args.options.userNames);
          if (invalidUpns !== true) {
            return `Invalid UPNs found for option 'userNames': ${invalidUpns}.`;
          }
        }

        if (args.options.subgroupIds !== undefined) {
          const invalidGuids = validation.isValidGuidArray(args.options.subgroupIds);
          if (invalidGuids !== true) {
            return `Invalid GUIDs found for option 'subgroupIds': ${invalidGuids}.`;
          }
        }

        if (args.options.role !== undefined && this.roleValues.indexOf(args.options.role) === -1) {
          return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }

        if ((args.options.subgroupIds !== undefined || args.options.subgroupNames !== undefined) && args.options.role?.toLowerCase() !== 'member') {
          return `When removing subgroups, the 'role' option must be set to 'Member'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupName'] },
      { options: ['userIds', 'userNames', 'subgroupIds', 'subgroupNames'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('groupId', 'groupName', 'ids', 'userNames', 'subgroupIds', 'subgroupNames', 'role');
    this.types.boolean.push('force', 'suppressNotFound');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const removeUsers = async (): Promise<void> => {
        if (this.verbose) {
          await logger.logToStderr(`Removing user(s) ${args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames} from group ${args.options.groupId || args.options.groupName}...`);
        }

        const groupId = await this.getGroupId(logger, args.options);
        const userIds = await this.getPrincipalIds(logger, args.options);

        const endpoints = [];
        if (!args.options.role || args.options.role === 'Owner') {
          endpoints.push(...userIds.map(id => `/groups/${groupId}/owners/${id}/$ref`));
        }
        if (!args.options.role || args.options.role === 'Member') {
          endpoints.push(...userIds.map(id => `/groups/${groupId}/members/${id}/$ref`));
        }

        for (let i = 0; i < endpoints.length; i += 20) {
          const endpointsBatch = endpoints.slice(i, i + 20);
          const requestOptions: CliRequestOptions = {
            url: `${this.resource}/v1.0/$batch`,
            headers: {
              'content-type': 'application/json;odata.metadata=none'
            },
            responseType: 'json',
            data: {
              requests: endpointsBatch.map((ep, index) => ({
                id: index + 1,
                method: 'DELETE',
                url: ep,
                headers: {
                  'content-type': 'application/json;odata.metadata=none'
                }
              }))
            } as GraphBatchRequest
          };

          const res = await request.post<GraphBatchRequestResponse>(requestOptions);
          for (const response of res.responses) {
            // Suppress 404 errors if suppressNotFound is set
            if (response.status !== 204 && (!args.options.suppressNotFound || response.status !== 404)) {
              throw response.body;
            }
          }
        }
      };

      if (args.options.force) {
        await removeUsers();
      }
      else {
        const principals = args.options.userIds || args.options.userNames || args.options.subgroupIds || args.options.subgroupNames;
        const principalsList = principals!.split(',');
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${principalsList.length} principal(s) from group '${args.options.groupId || args.options.groupName}'?` });

        if (result) {
          await removeUsers();
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(logger: Logger, options: Options): Promise<string> {
    if (options.groupId) {
      return options.groupId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving ID of group '${options.groupName}'...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupName!);
  }

  private async getPrincipalIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.userIds) {
      return options.userIds.split(',').map(i => i.trim());
    }

    if (options.subgroupIds) {
      return options.subgroupIds.split(',').map(i => i.trim());
    }

    if (options.userNames) {
      if (this.verbose) {
        await logger.logToStderr('Retrieving ID(s) of user(s)...');
      }

      return entraUser.getUserIdsByUpns(options.userNames!.split(',').map(u => u.trim()));
    }

    // Subgroup names were specified
    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of subgroup(s)...');
    }

    const subGroupIds: string[] = [];
    for (const subgroupName of options.subgroupNames!.split(',')) {
      const groupId = await entraGroup.getGroupIdByDisplayName(subgroupName.trim());
      subGroupIds.push(groupId);
    }
    return subGroupIds;
  }
}

export default new EntraGroupMemberRemoveCommand();