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
  confirm?: boolean;
}

class SpoListItemRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from list item permissions';
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
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--listItemId <listItemId>'
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

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        if (args.options.listItemId && isNaN(args.options.listItemId)) {
          return `Specified listItemId ${args.options.listItemId} is not a number`;
        }

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
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

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeRoleAssignment: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing role assignment from listitem in site at ${args.options.webUrl}...`);
      }

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

      if (args.options.upn) {
        this.GetUserPrincipalId(args.options)
          .then((userPrincipalId: number) => {
            args.options.principalId = userPrincipalId;
            this.RemoveRoleAssignment(requestUrl, logger, args.options, cb);
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else if (args.options.groupName) {
        this.GetGroupPrincipalId(args.options)
          .then((groupPrincipalId: number) => {
            args.options.principalId = groupPrincipalId;
            this.RemoveRoleAssignment(requestUrl, logger, args.options, cb);
          }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
      }
      else {
        this.RemoveRoleAssignment(requestUrl, logger, args.options, cb);
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
        message: `Are you sure you want to remove role assignment from listitem ${args.options.listItemId} from list ${args.options.listId || args.options.listTitle} from site ${args.options.webUrl}?`
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

  private RemoveRoleAssignment(requestUrl: string, logger: Logger, options: Options, cb: () => void): void {
    const requestOptions: any = {
      url: `${requestUrl}roleassignments/removeroleassignment(principalid='${options.principalId}')`,
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

module.exports = new SpoListItemRoleAssignmentRemoveCommand();