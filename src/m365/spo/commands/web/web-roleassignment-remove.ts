import { Cli, CommandOutput, Logger } from '../../../../cli';
import Command, { CommandErrorWithOutput } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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
  principalId?: number;
  upn?: string;
  groupName?: string;
  confirm?: boolean;
}

class SpoWebRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from web permissions';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeRoleAssignment: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing role assignment from web ${args.options.webUrl}...`);
      }

      if (args.options.upn) {
        this.GetUserPrincipalId(args.options)
          .then((userPrincipalId: number) => {
            args.options.principalId = userPrincipalId;
            this.RemoveRoleAssignment(logger, args.options, cb);
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else if (args.options.groupName) {
        this.GetGroupPrincipalId(args.options)
          .then((groupPrincipalId: number) => {
            args.options.principalId = groupPrincipalId;
            this.RemoveRoleAssignment(logger, args.options, cb);
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else {
        this.RemoveRoleAssignment(logger, args.options, cb);
      }
    };

    if (args.options.confirm) {
      removeRoleAssignment();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove role assignment from web ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeRoleAssignment();
        }
      });
    }
  }

  private RemoveRoleAssignment(logger: Logger, options: Options, cb: () => void): void {
    const requestOptions: any = {
      url: `${options.webUrl}/_api/web/roleassignments/removeroleassignment(principalid='${options.principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

module.exports = new SpoWebRoleAssignmentRemoveCommand();