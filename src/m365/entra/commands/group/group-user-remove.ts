import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { GraphBatchRequest, GraphBatchRequestResponse } from '../../../../utils/types.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  ids?: string;
  userNames?: string;
  role?: string;
  suppressNotFound?: boolean;
  force?: boolean;
}

class EntraGroupUserRemoveCommand extends GraphCommand {
  private readonly roleValues = ['Owner', 'Member'];

  public get name(): string {
    return commands.GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes users from a Microsoft Entra ID group';
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
        groupDisplayName: typeof args.options.groupDisplayName !== 'undefined',
        ids: typeof args.options.ids !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined',
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
        option: '-n, --groupDisplayName [groupDisplayName]'
      },
      {
        option: '--ids [ids]'
      },
      {
        option: '--userNames [userNames]'
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

        if (args.options.ids !== undefined) {
          const ids = args.options.ids.split(',').map(i => i.trim());
          if (!validation.isValidGuidArray(ids)) {
            const invalidGuid = ids.find(id => !validation.isValidGuid(id));
            return `'${invalidGuid}' is not a valid GUID for option 'ids'.`;
          }
        }

        if (args.options.userNames !== undefined) {
          const isValidUserPrincipalNameArray = validation.isValidUserPrincipalNameArray(args.options.userNames.split(',').map(u => u.trim()));
          if (isValidUserPrincipalNameArray !== true) {
            return `User principal name '${isValidUserPrincipalNameArray}' is invalid for option 'userNames'.`;
          }
        }

        if (args.options.role !== undefined && this.roleValues.indexOf(args.options.role) === -1) {
          return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupDisplayName'] },
      { options: ['ids', 'userNames'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('groupId', 'groupDisplayName', 'ids', 'userNames', 'role');
    this.types.boolean.push('force', 'suppressNotFound');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const removeUsers = async (): Promise<void> => {
        if (this.verbose) {
          await logger.logToStderr(`Removing user(s) ${args.options.ids || args.options.userNames} from group ${args.options.groupId || args.options.groupDisplayName}...`);
        }

        const groupId = await this.getGroupId(logger, args.options);
        const userIds = await this.getUserIds(logger, args.options);

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
        const users = args.options.ids || args.options.userNames;
        const userList = users!.split(',');
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove ${userList.length} user(s) from group '${args.options.groupId || args.options.groupDisplayName}'?` });

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
      await logger.logToStderr(`Retrieving ID of group '${options.groupDisplayName}'...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupDisplayName!);
  }

  private async getUserIds(logger: Logger, options: Options): Promise<string[]> {
    if (options.ids) {
      return options.ids.split(',').map(i => i.trim());
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving ID(s) of user(s)...');
    }

    return entraUser.getUserIdsByUpns(options.userNames!.split(',').map(u => u.trim()));
  }
}

export default new EntraGroupUserRemoveCommand();