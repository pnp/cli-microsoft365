import { cli } from '../../../../cli/cli.js';
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
import spoUserGetCommand, { Options as SpoUserGetCommandOptions } from '../user/user-get.js';

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
  force?: boolean;
}

class SpoListRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from list permissions';
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
        force: (!(!args.options.force)).toString()
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
        option: '-f, --force'
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['principalId', 'upn', 'groupName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing role assignment from list in site at ${args.options.webUrl}...`);
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

        if (args.options.upn) {
          args.options.principalId = await this.getUserPrincipalId(args.options);
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
        else if (args.options.groupName) {
          args.options.principalId = await this.getGroupPrincipalId(args.options);
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
        else {
          await this.removeRoleAssignment(requestUrl, logger, args.options);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from list ${args.options.listId || args.options.listTitle} from site ${args.options.webUrl}?` });

      if (result) {
        await removeRoleAssignment();
      }
    }
  }

  private async removeRoleAssignment(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}roleassignments/removeroleassignment(principalid='${options.principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getGroupPrincipalId(options: Options): Promise<number> {
    const groupGetCommandOptions: SpoGroupGetCommandOptions = {
      webUrl: options.webUrl,
      name: options.groupName,
      output: 'json',
      debug: this.debug,
      verbose: this.verbose
    };

    const output = await cli.executeCommandWithOutput(spoGroupGetCommand as Command, { options: { ...groupGetCommandOptions, _: [] } });
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

    const output = await cli.executeCommandWithOutput(spoUserGetCommand as Command, { options: { ...userGetCommandOptions, _: [] } });

    const getUserOutput = JSON.parse(output.stdout);
    return getUserOutput.Id;
  }
}

export default new SpoListRoleAssignmentRemoveCommand();