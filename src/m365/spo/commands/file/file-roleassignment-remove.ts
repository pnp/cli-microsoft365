import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { urlUtil } from '../../../../utils/urlUtil';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoFileGetCommand from './file-get';
import { Options as SpoFileGetCommandOptions } from './file-get';
import * as SpoUserGetCommand from '../user/user-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import { Options as SpoGroupGetCommandOptions } from '../group/group-get';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  fileUrl?: string;
  fileId?: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  confirm?: boolean;
}

class SpoFileRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from a file.';
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
        confirm: !!args.options.confirm
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
        option: '-i, --fileId [fileId]'
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
        option: '--confirm'
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

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.fileId && !validation.isValidGuid(args.options.fileId)) {
          return `${args.options.fileId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['fileUrl', 'fileId'] },
      { options: ['upn', 'groupName', 'principalId'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment: () => Promise<void> = async (): Promise<void> => {
      try {
        const fileURL: string = await this.getFileURL(args);

        let principalId: number;
        if (args.options.groupName) {
          principalId = await this.getGroupPrincipalId(args.options);
        }
        else if (args.options.upn) {
          principalId = await this.getUserPrincipalId(args.options);
        }
        else {
          principalId = args.options.principalId!;
        }

        const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, fileURL);
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativeUrl('${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeRoleAssignment();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove role assignment from file ${args.options.fileUrl || args.options.fileId} from site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await removeRoleAssignment();
      }
    }
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

    const output = await Cli.executeCommandWithOutput(SpoFileGetCommand as Command, { options: { ...options, _: [] } });
    const getFileOutput = JSON.parse(output.stdout);
    return getFileOutput.ServerRelativeUrl;
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
}

module.exports = new SpoFileRoleAssignmentRemoveCommand();