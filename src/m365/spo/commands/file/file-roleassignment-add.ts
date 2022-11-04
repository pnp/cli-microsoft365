import { Cli, CommandOutput } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import * as SpoUserGetCommand from '../user/user-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import { Options as SpoGroupGetCommandOptions } from '../group/group-get';
import * as SpoRoleDefinitionListCommand from '../roledefinition/roledefinition-list';
import { Options as SpoRoleDefinitionListCommandOptions } from '../roledefinition/roledefinition-list';
import { RoleDefinition } from '../roledefinition/RoleDefinition';
import * as SpoFileGetCommand from './file-get';
import { Options as SpoFileGetCommandOptions } from './file-get';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import { formatting } from '../../../../utils/formatting';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  roleDefinitionId?: number;
  roleDefinitionName?: string;
}

class SpoFileRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to the specified file.';
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
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined',
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
        option: '--fileUrl [fileUrl]'
      },
      {
        option: 'i, --fileId [fileId]'
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

        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
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
      ['fileId', 'fileUrl'],
      ['principalId', 'upn', 'groupName'],
      ['roleDefinitionId', 'roleDefinitionName']
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding role assignment to file in site at ${args.options.webUrl}...`);
    }

    try {
      const fileUrl: string = await this.getFileURL(args);
      const roleDefinitionId = await this.getRoleDefinitionId(args.options);
      if (args.options.upn) {
        const upnPrincipalId = await this.getUserPrincipalId(args.options);
        await this.addRoleAssignment(fileUrl, args.options.webUrl, upnPrincipalId, roleDefinitionId);
      }
      else if (args.options.groupName) {
        const groupPrincipalId = await this.getGroupPrincipalId(args.options);
        await this.addRoleAssignment(fileUrl, args.options.webUrl, groupPrincipalId, roleDefinitionId);
      }
      else {
        await this.addRoleAssignment(fileUrl, args.options.webUrl, args.options.principalId!, roleDefinitionId);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(fileUrl: string, webUrl: string, principalId: number, roleDefinitionId: number): Promise<void> {
    try {
      const requestOptions: any = {
        url: `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
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

  private async getFileURL(args: CommandArgs): Promise<string> {
    if (args.options.fileUrl) {
      return args.options.fileUrl;
    }

    const options: SpoFileGetCommandOptions = {
      webUrl: args.options.webUrl,
      id: args.options.fileId,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(SpoFileGetCommand as Command, { options: { ...options, _: [] } });
    const getFileOutput = JSON.parse(output.stdout);
    return getFileOutput.ServerRelativeUrl;
  }
}

module.exports = new SpoFileRoleAssignmentAddCommand();
