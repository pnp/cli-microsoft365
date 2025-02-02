import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId?: string;
  groupDisplayName?: string;
  groupName?: string;
  ids?: string;
  userNames?: string;
  role: string;
}

class EntraGroupMemberAddCommand extends GraphCommand {
  private readonly roleValues = ['Owner', 'Member'];

  public get name(): string {
    return commands.GROUP_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a member to a Microsoft Entra ID group';
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
        groupName: typeof args.options.groupName !== 'undefined',
        ids: typeof args.options.ids !== 'undefined',
        userNames: typeof args.options.userNames !== 'undefined'
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
        option: "--groupName [groupName]"
      },
      {
        option: '--ids [ids]'
      },
      {
        option: '--userNames [userNames]'
      },
      {
        option: '-r, --role <role>',
        autocomplete: this.roleValues
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID for option groupId.`;
        }

        if (args.options.ids) {
          const isValidGUIDArrayResult = validation.isValidGuidArray(args.options.ids);
          if (isValidGUIDArrayResult !== true) {
            return `The following GUIDs are invalid for the option 'ids': ${isValidGUIDArrayResult}.`;
          }
        }

        if (args.options.userNames) {
          const isValidUPNArrayResult = validation.isValidUserPrincipalNameArray(args.options.userNames);
          if (isValidUPNArrayResult !== true) {
            return `The following user principal names are invalid for the option 'userNames': ${isValidUPNArrayResult}.`;
          }
        }

        if (this.roleValues.indexOf(args.options.role) === -1) {
          return `Option 'role' must be one of the following values: ${this.roleValues.join(', ')}.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['groupId', 'groupDisplayName', 'groupName'] },
      { options: ['ids', 'userNames'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('groupId', 'groupDisplayName', 'groupName', 'ids', 'userNames', 'role');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.groupDisplayName) {
        await this.warn(logger, `Option 'groupDisplayName' is deprecated and will be removed in the next major release.`);
      }

      if (this.verbose) {
        await logger.logToStderr(`Adding member(s) ${args.options.ids || args.options.userNames} to group ${args.options.groupId || args.options.groupDisplayName || args.options.groupName}...`);
      }

      const groupId = await this.getGroupId(logger, args.options);
      const userIds = await this.getUserIds(logger, args.options);

      for (let i = 0; i < userIds.length; i += 400) {
        const userIdsBatch = userIds.slice(i, i + 400);
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/$batch`,
          headers: {
            'content-type': 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {
            requests: []
          }
        };

        for (let j = 0; j < userIdsBatch.length; j += 20) {
          const userIdsChunk = userIdsBatch.slice(j, j + 20);
          requestOptions.data.requests.push({
            id: j + 1,
            method: 'PATCH',
            url: `/groups/${groupId}`,
            headers: {
              'content-type': 'application/json;odata.metadata=none'
            },
            body: {
              [`${args.options.role === 'Member' ? 'members' : 'owners'}@odata.bind`]: userIdsChunk.map(u => `${this.resource}/v1.0/directoryObjects/${u}`)
            }
          });
        }

        const res = await request.post<{ responses: { status: number; body: any }[] }>(requestOptions);
        for (const response of res.responses) {
          if (response.status !== 204) {
            throw response.body;
          }
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
      await logger.logToStderr(`Retrieving ID of group ${options.groupDisplayName || options.groupName}...`);
    }

    return entraGroup.getGroupIdByDisplayName(options.groupDisplayName! || options.groupName!);
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

export default new EntraGroupMemberAddCommand();