import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
  roleDefinitionId?: number;
  roleDefinitionName?: string;
}

class SpoListRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to list permissions';
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
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined',
        roleDefinitionId: typeof args.options.roleDefinitionId !== 'undefined',
        roleDefinitionName: typeof args.options.roleDefinitionName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--principalId [principalId]'
      },
      {
        option: '--upn [upn]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
      },
      {
        option: '--roleDefinitionId [roleDefinitionId]'
      },
      {
        option: '--roleDefinitionName [roleDefinitionName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `'${args.options.listId}' is not a valid GUID for option listId.`;
        }

        if (args.options.upn && !validation.isValidUserPrincipalName(args.options.upn)) {
          return `'${args.options.upn}' is not a valid user principal name for option upn.`;
        }

        if (args.options.principalId && !validation.isValidPositiveInteger(args.options.principalId)) {
          return `Specified principalId '${args.options.principalId}' is not a valid number.`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }

        if (args.options.roleDefinitionId && !validation.isValidPositiveInteger(args.options.roleDefinitionId)) {
          return `Specified roleDefinitionId '${args.options.roleDefinitionId}' is not a valid number.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] },
      { options: ['roleDefinitionId', 'roleDefinitionName'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'listId', 'listTitle', 'listUrl', 'upn', 'groupName', 'entraGroupId', 'entraGroupName', 'roleDefinitionName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to list in site at ${args.options.webUrl}...`);
    }

    try {
      let requestUrl: string = `${args.options.webUrl}/_api/web/`;
      if (args.options.listId) {
        requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
      }
      else if (args.options.listTitle) {
        requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
      }

      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      let principalId: number | undefined = args.options.principalId;
      if (args.options.upn) {
        const user = await spo.ensureUser(args.options.webUrl, args.options.upn);
        principalId = user.Id;
      }
      else if (args.options.groupName) {
        const group = await spo.getGroupByName(args.options.webUrl, args.options.groupName, logger, this.verbose);
        principalId = group.Id;
      }
      else if (args.options.entraGroupId || args.options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        let group: Group;
        if (args.options.entraGroupId) {
          group = await entraGroup.getGroupById(args.options.entraGroupId);
        }
        else {
          group = await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);
        }

        const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
        principalId = siteUser.Id;
      }

      await this.addRoleAssignment(requestUrl, principalId!, roleDefinitionId);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(requestUrl: string, principalId: number, roleDefinitionId: number): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<number> {
    if (options.roleDefinitionId) {
      return options.roleDefinitionId;
    }

    const roleDefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName!, logger, this.verbose);
    return roleDefinition.Id;
  }
}

export default new SpoListRoleAssignmentAddCommand();