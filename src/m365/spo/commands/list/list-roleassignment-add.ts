import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoUserGetCommand from '../user/user-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import { Options as SpoGroupGetCommandOptions } from '../group/group-get';
import * as SpoRoleDefinitionListCommand from '../roledefinition/roledefinition-list';
import { Options as SpoRoleDefinitionListCommandOptions } from '../roledefinition/roledefinition-list';
import { RoleDefinition } from '../roledefinition/RoleDefinition';

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
          return `${args.options.listId} is not a valid GUID`;
        }

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
          return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['principalId', 'upn', 'groupName'] },
      { options: ['roleDefinitionId', 'roleDefinitionName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding role assignment to list in site at ${args.options.webUrl}...`);
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

      args.options.roleDefinitionId = await this.getRoleDefinitionId(args.options);
      if (args.options.upn) {
        args.options.principalId = await this.getUserPrincipalId(args.options);
        await this.addRoleAssignment(requestUrl, logger, args.options);
      }
      else if (args.options.groupName) {
        args.options.principalId = await this.getGroupPrincipalId(args.options);
        await this.addRoleAssignment(requestUrl, logger, args.options);
      }
      else {
        await this.addRoleAssignment(requestUrl, logger, args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}roleassignments/addroleassignment(principalid='${options.principalId}',roledefid='${options.roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options): Promise<number> {
    if (!options.roleDefinitionName) {
      return options.roleDefinitionId as number;
    }

    const roleDefinitionListCommandOptions: SpoRoleDefinitionListCommandOptions = {
      webUrl: options.webUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoRoleDefinitionListCommand as Command, { options: { ...roleDefinitionListCommandOptions, _: [] } });
    const getRoleDefinitionListOutput = JSON.parse(output.stdout);
    const roleDefinitionId: number = getRoleDefinitionListOutput.find((role: RoleDefinition) => role.Name === options.roleDefinitionName).Id;
    return roleDefinitionId;
  }

  private async getGroupPrincipalId(options: Options): Promise<number> {
    const groupGetCommandOptions: SpoGroupGetCommandOptions = {
      webUrl: options.webUrl,
      name: options.groupName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } });
    const getGroupOutput = JSON.parse(output.stdout);
    return getGroupOutput.Id;
  }

  private async getUserPrincipalId(options: Options): Promise<number> {
    const userGetCommandOptions: SpoUserGetCommandOptions = {
      webUrl: options.webUrl,
      email: options.upn,
      id: undefined,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.Id;
  }
}

module.exports = new SpoListRoleAssignmentAddCommand();