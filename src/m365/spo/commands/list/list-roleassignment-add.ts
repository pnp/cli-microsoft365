import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandErrorWithOutput } from '../../../../Command';
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

        const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
        if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either list id or title or list url`;
        }

        if (listOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least list id or title or list url`;
        }

        const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName];
        if (principalOptions.some(item => item !== undefined) && principalOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either principalId id or upn or groupName`;
        }

        if (principalOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least principalId id or upn or groupName`;
        }

        const roleDefinitionOptions: any[] = [args.options.roleDefinitionId, args.options.roleDefinitionName];
        if (roleDefinitionOptions.some(item => item !== undefined) && roleDefinitionOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either roleDefinitionId id or roleDefinitionName`;
        }

        if (roleDefinitionOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least roleDefinitionId id or roleDefinitionName`;
        }

        return true;
      }
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

      args.options.roleDefinitionId = await this.GetRoleDefinitionId(args.options);
      if (args.options.upn) {
        args.options.principalId = await this.GetUserPrincipalId(args.options);
        await this.AddRoleAssignment(requestUrl, logger, args.options);
      }
      else if (args.options.groupName) {
        args.options.principalId = await this.GetGroupPrincipalId(args.options);
        this.AddRoleAssignment(requestUrl, logger, args.options);
      }
      else {
        await this.AddRoleAssignment(requestUrl, logger, args.options);
      }
      
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private AddRoleAssignment(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: any = {
      url: `${requestUrl}roleassignments/addroleassignment(principalid='${options.principalId}',roledefid='${options.roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request
      .post(requestOptions)
      .then(_ => Promise.resolve())
      .catch((err: any) => Promise.reject(err));
  }

  private GetRoleDefinitionId(options: Options): Promise<number> {
    if (!options.roleDefinitionName) {
      return Promise.resolve(options.roleDefinitionId as number);
    }

    const roleDefinitionListCommandOptions: SpoRoleDefinitionListCommandOptions = {
      webUrl: options.webUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(SpoRoleDefinitionListCommand as Command, { options: { ...roleDefinitionListCommandOptions, _: [] } })
      .then((output: CommandOutput): Promise<number> => {
        const getRoleDefinitionListOutput = JSON.parse(output.stdout);
        const roleDefinitionId: number = getRoleDefinitionListOutput.find((role: RoleDefinition) => role.Name === options.roleDefinitionName).Id;
        return Promise.resolve(roleDefinitionId);
      }, (err: CommandErrorWithOutput) => {
        return Promise.reject(err);
      });
  }

  private GetGroupPrincipalId(options: Options): Promise<number> {
    const groupGetCommandOptions: SpoGroupGetCommandOptions = {
      webUrl: options.webUrl,
      name: options.groupName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(SpoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } })
      .then((output: CommandOutput): Promise<number> => {
        const getGroupOutput = JSON.parse(output.stdout);
        return Promise.resolve(getGroupOutput.Id);
      }, (err: CommandErrorWithOutput) => {
        return Promise.reject(err);
      });
  }

  private GetUserPrincipalId(options: Options): Promise<number> {
    const userGetCommandOptions: SpoUserGetCommandOptions = {
      webUrl: options.webUrl,
      email: options.upn,
      id: undefined,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    return Cli.executeCommandWithOutput(SpoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } })
      .then((output: CommandOutput): Promise<number> => {
        const getUserOutput = JSON.parse(output.stdout);
        return Promise.resolve(getUserOutput.Id);
      }, (err: CommandErrorWithOutput) => {
        return Promise.reject(err);
      });
  }
}

module.exports = new SpoListRoleAssignmentAddCommand();