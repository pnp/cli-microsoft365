import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
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
  listItemId: number;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  roleDefinitionId?: number;
  roleDefinitionName?: string;
}

class SpoListItemRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to a listitem.';
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
        option: '--listItemId <listItemId>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
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

        if (args.options.listItemId && isNaN(args.options.listItemId)) {
          return `Specified listItemId ${args.options.listItemId} is not a number`;
        }

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
          return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }

        const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either list id or title or list url`;
        }

        if (listOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least list id or title or list url`;
        }

        const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName];
        if (!principalOptions.some(item => item !== undefined)) {
          return `Specify either principalId, upn or groupName`;
        }

        if (principalOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either principalId, upn or groupName but not multiple`;
        }

        const roleDefinitionOptions: any[] = [args.options.roleDefinitionId, args.options.roleDefinitionName];
        if (!roleDefinitionOptions.some(item => item !== undefined)) {
          return `Specify either roleDefinitionId or roleDefinitionName`;
        }

        if (roleDefinitionOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either roleDefinitionId or roleDefinitionName but not multiple`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding role assignment to listitem in site at ${args.options.webUrl}...`);
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

      requestUrl += `items(${args.options.listItemId})/`;

      args.options.roleDefinitionId = await this.getRoleDefinitionId(args.options);
      if (args.options.upn) {
        args.options.principalId = await this.getUserPrincipalId(args.options);
        await this.addRoleAssignment(requestUrl, args.options);
      }
      else if (args.options.groupName) {
        args.options.principalId = await this.getGroupPrincipalId(args.options);
        await this.addRoleAssignment(requestUrl, args.options);
      }
      else {
        await this.addRoleAssignment(requestUrl, args.options);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(requestUrl: string, options: Options): Promise<void> {
    try {
      const requestOptions: any = {
        url: `${requestUrl}roleassignments/addroleassignment(principalid='${options.principalId}',roledefid='${options.roleDefinitionId}')`,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      return Promise.reject(err);
    }
  }

  private async getRoleDefinitionId(options: Options): Promise<number> {
    if (!options.roleDefinitionName) {
      return options.roleDefinitionId!;
    }

    try {
      const roleDefinitionListCommandOptions: SpoRoleDefinitionListCommandOptions = {
        webUrl: options.webUrl,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const output: CommandOutput = await Cli.executeCommandWithOutput(SpoRoleDefinitionListCommand as Command, { options: { ...roleDefinitionListCommandOptions, _: [] } });
      const getRoleDefinitionListOutput = JSON.parse(output.stdout);
      const roleDefinitionId: number = getRoleDefinitionListOutput.find((role: RoleDefinition) => role.Name === options.roleDefinitionName).Id;
      return roleDefinitionId;
    }
    catch (err: any) {
      return Promise.reject(err);
    }
  }

  private async getGroupPrincipalId(options: Options): Promise<number> {
    try {
      const groupGetCommandOptions: SpoGroupGetCommandOptions = {
        webUrl: options.webUrl,
        name: options.groupName,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };

      const output: CommandOutput = await Cli.executeCommandWithOutput(SpoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } });
      const getGroupOutput = JSON.parse(output.stdout);
      return getGroupOutput.Id;
    }
    catch (err: any) {
      return Promise.reject(err);
    }
  }

  private async getUserPrincipalId(options: Options): Promise<number> {
    try {
      const userGetCommandOptions: SpoUserGetCommandOptions = {
        webUrl: options.webUrl,
        email: options.upn,
        id: undefined,
        output: 'json',
        debug: this.debug,
        verbose: this.verbose
      };
  
      const output: CommandOutput = await Cli.executeCommandWithOutput(SpoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } });
      const getUserOutput = JSON.parse(output.stdout);
      return getUserOutput.Id;
    }
    catch (err: any) {
      return Promise.reject(err);
    }
  }
}

module.exports = new SpoListItemRoleAssignmentAddCommand();