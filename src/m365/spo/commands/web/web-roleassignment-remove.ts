import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { cli } from '../../../../cli/cli.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
  force?: boolean;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined',
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
        option: '--principalId [principalId]'
      },
      {
        option: '--upn [upn]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--entraGroupId [entraGroupId]'
      },
      {
        option: '--entraGroupName [entraGroupName]'
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

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoleAssignment(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from web ${args.options.webUrl}?` });

      if (result) {
        await this.removeRoleAssignment(logger, args.options);
      }
    }
  }

  private async removeRoleAssignment(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing role assignment from web ${options.webUrl}...`);
    }

    try {
      if (options.upn) {
        const principalId = await this.getUserPrincipalId(options, logger);
        await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
      }
      else if (options.groupName) {
        const principalId = await this.getGroupPrincipalId(options, logger);
        await this.removeRoleAssignmentWithOptions(options.webUrl, principalId, logger);
      }
      else if (options.entraGroupId || options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = options.entraGroupId
          ? await entraGroup.getGroupById(options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
        await this.removeRoleAssignmentWithOptions(options.webUrl, siteUser.Id, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeRoleAssignmentWithOptions(webUrl: string, principalId: number, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing role assignment...');
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/roleassignments/removeroleassignment(principalid='${principalId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }
}

export default new SpoWebRoleAssignmentRemoveCommand();