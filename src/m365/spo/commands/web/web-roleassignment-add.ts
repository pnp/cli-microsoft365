import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
  roleDefinitionId?: number;
  roleDefinitionName?: string;
}

class SpoWebRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to web';
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

        if (args.options.principalId && isNaN(args.options.principalId)) {
          return `Specified principalId ${args.options.principalId} is not a number`;
        }

        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
          return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
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
      { options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName'] },
      { options: ['roleDefinitionId', 'roleDefinitionName'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to web ${args.options.webUrl}...`);
    }

    try {
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);

      if (args.options.upn) {
        const principalId = await this.getUserPrincipalId(args.options, logger);
        await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
      }
      else if (args.options.groupName) {
        const principalId = await this.getGroupPrincipalId(args.options, logger);
        await this.addRoleAssignment(args.options.webUrl, principalId, roleDefinitionId, logger);
      }
      else if (args.options.entraGroupId || args.options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = args.options.entraGroupId
          ? await entraGroup.getGroupById(args.options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
        await this.addRoleAssignment(args.options.webUrl, siteUser.Id, roleDefinitionId, logger);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(webUrl: string, principalId: number, roleDefinitionId: number, logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Adding role assignment...');
    }

    const requestOptions: any = {
      url: `${webUrl}/_api/web/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    await request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<number> {
    if (!options.roleDefinitionName) {
      return options.roleDefinitionId as number;
    }

    const roledefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);

    return roledefinition.Id;
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

export default new SpoWebRoleAssignmentAddCommand();