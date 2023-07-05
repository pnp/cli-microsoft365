import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  principalId?: number;
  upn?: string;
  groupName?: string;
  entraGroupId?: string;
  entraGroupName?: string;
  roleDefinitionId?: number;
  roleDefinitionName?: string;
}

class SpoFolderRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to the specified folder.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
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

        if (args.options.entraGroupId && !validation.isValidGuid(args.options.entraGroupId)) {
          return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId.`;
        }

        if (args.options.roleDefinitionId && isNaN(args.options.roleDefinitionId)) {
          return `Specified roleDefinitionId ${args.options.roleDefinitionId} is not a number`;
        }

        const principalOptions: any[] = [args.options.principalId, args.options.upn, args.options.groupName, args.options.entraGroupId, args.options.entraGroupName];
        if (!principalOptions.some(item => item !== undefined)) {
          return `Specify either principalId, upn, groupName, entraGroupId or entraGroupName`;
        }

        if (principalOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either principalId, upn, groupName, entraGroupId or entraGroupName but not multiple`;
        }

        const roleDefinitionOptions: any[] = [args.options.roleDefinitionId, args.options.roleDefinitionName];
        if (!roleDefinitionOptions.some(item => item !== undefined)) {
          return `Specify either roleDefinitionId id or roleDefinitionName`;
        }

        if (roleDefinitionOptions.filter(item => item !== undefined).length > 1) {
          return `Specify either roleDefinitionId id or roleDefinitionName but not both`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl', 'upn', 'groupName', 'roleDefinitionName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to folder in site at ${args.options.webUrl}...`);
    }

    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
    const roleFolderUrl: string = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);

    try {
      let requestUrl: string = `${args.options.webUrl}/_api/web/`;
      if (roleFolderUrl.split('/').length === 2) {
        requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
      }
      else {
        requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
      }

      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      let principalId: number | undefined = args.options.principalId;
      if (args.options.upn) {
        principalId = await this.getUserPrincipalId(args.options, logger);
      }
      else if (args.options.groupName) {
        principalId = await this.getGroupPrincipalId(args.options, logger);
      }
      else if (args.options.entraGroupId || args.options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = args.options.entraGroupId
          ? await entraGroup.getGroupById(args.options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(args.options.webUrl, group);
        principalId = siteUser.Id;
      }

      await this.addRoleAssignment(requestUrl, principalId!, roleDefinitionId);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async addRoleAssignment(requestUrl: string, principalId: number, roleDefinitionId: number): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<number> {
    if (!options.roleDefinitionName) {
      return options.roleDefinitionId as number;
    }

    const roleDefintion = await spo.getRoleDefintionByName(options.webUrl, options.roleDefinitionName!, logger, this.verbose);
    return roleDefintion.Id;
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id as number;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id as number;
  }
}

export default new SpoFolderRoleAssignmentAddCommand();