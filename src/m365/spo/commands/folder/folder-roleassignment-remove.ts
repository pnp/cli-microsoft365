import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandErrorWithOutput } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as SpoUserGetCommand from '../user/user-get';
import { Options as SpoUserGetCommandOptions } from '../user/user-get';
import * as SpoGroupGetCommand from '../group/group-get';
import { Options as SpoGroupGetCommandOptions } from '../group/group-get';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  confirm?: boolean;
}

class SpoFolderRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from folder permissions';
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
        folderUrl: typeof args.options.folderUrl !== 'undefined',
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderUrl <folderUrl>'
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

        const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName];
        if (principalOptions.some(item => item !== undefined) && principalOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either principalId id or upn or groupName`;
        }

        if (principalOptions.filter(item => item !== undefined).length === 0) {
          return `Specify at least principalId id or upn or groupName`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment: () => Promise<void> = async (): Promise<void> => {
      if (this.verbose) {
        logger.logToStderr(`Removing role assignment from folder in site at ${args.options.webUrl}...`);
      }
      const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
      const requestUrl: string = `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/ListItemAllFields`;

      try { 
        if (args.options.upn) {
          args.options.principalId = await this.GetUserPrincipalId(args.options);
          await this.RemoveRoleAssignment(requestUrl, logger, args.options);
        }
        else if (args.options.groupName) {
          args.options.principalId = await this.GetGroupPrincipalId(args.options);
          await this.RemoveRoleAssignment(requestUrl, logger, args.options);
        }
        else {
          await this.RemoveRoleAssignment(requestUrl, logger, args.options);
        }
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
        message: `Are you sure you want to remove role assignment from folder ${args.options.folderUrl} from site ${args.options.webUrl}?`
      });
      
      if (result.continue) {
        await removeRoleAssignment();
      }
    }
  }

  private RemoveRoleAssignment(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: any = {
      url: `${requestUrl}/roleassignments/removeroleassignment(principalid='${options.principalId}')`,
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

module.exports = new SpoFolderRoleAssignmentRemoveCommand();