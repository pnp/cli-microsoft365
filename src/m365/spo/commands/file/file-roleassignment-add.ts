import { Cli, CommandOutput } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import Command from '../../../../Command.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import spoGroupGetCommand, { Options as SpoGroupGetCommandOptions } from '../group/group-get.js';
import spoRoleDefinitionListCommand, { Options as SpoRoleDefinitionListCommandOptions } from '../roledefinition/roledefinition-list.js';
import { RoleDefinition } from '../roledefinition/RoleDefinition.js';
import spoUserGetCommand, { Options as SpoUserGetCommandOptions } from '../user/user-get.js';
import spoFileGetCommand, { Options as SpoFileGetCommandOptions } from './file-get.js';

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
      { options: ['fileId', 'fileUrl'] },
      { options: ['principalId', 'upn', 'groupName'] },
      { options: ['roleDefinitionId', 'roleDefinitionName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to file in site at ${args.options.webUrl}...`);
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
    const requestOptions: CliRequestOptions = {
      url: `${webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
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
      return options.roleDefinitionId!;
    }

    const roleDefinitionListCommandOptions: SpoRoleDefinitionListCommandOptions = {
      webUrl: options.webUrl,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output: CommandOutput = await Cli.executeCommandWithOutput(spoRoleDefinitionListCommand as Command, { options: { ...roleDefinitionListCommandOptions, _: [] } });
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

    const output: CommandOutput = await Cli.executeCommandWithOutput(spoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } });
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

    const output: CommandOutput = await Cli.executeCommandWithOutput(spoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } });
    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.Id;
  }

  private async getFileURL(args: CommandArgs): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const options: SpoFileGetCommandOptions = {
      webUrl: args.options.webUrl,
      id: args.options.fileId,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await Cli.executeCommandWithOutput(spoFileGetCommand as Command, { options: { ...options, _: [] } });
    const getFileOutput = JSON.parse(output.stdout);
    return getFileOutput.ServerRelativeUrl;
  }
}

export default new SpoFileRoleAssignmentAddCommand();
