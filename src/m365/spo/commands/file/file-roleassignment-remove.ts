import { Group } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { FileProperties } from './FileProperties.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

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
  entraGroupId?: string;
  entraGroupName?: string;
  force?: boolean;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        fileUrl: typeof args.options.fileUrl !== 'undefined',
        fileId: typeof args.options.fileId !== 'undefined',
        principalId: typeof args.options.principalId !== 'undefined',
        upn: typeof args.options.upn !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        entraGroupId: typeof args.options.entraGroupId !== 'undefined',
        entraGroupName: typeof args.options.entraGroupName !== 'undefined',
        force: !!args.options.force
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
          return `'${args.options.entraGroupId}' is not a valid GUID for option entraGroupId`;
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
      { options: ['upn', 'groupName', 'principalId', 'entraGroupId', 'entraGroupName'] }
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'fileUrl', 'fileId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing role assignment for ${args.options.groupName || args.options.upn} from file ${args.options.fileUrl || args.options.fileId}`);
      }

      try {
        const fileURL: string = await this.getFileURL(args, logger);

        let principalId: number;
        if (args.options.groupName) {
          principalId = await this.getGroupPrincipalId(args.options, logger);
        }
        else if (args.options.upn) {
          principalId = await this.getUserPrincipalId(args.options, logger);
        }
        else if (args.options.entraGroupId || args.options.entraGroupName) {
          if (this.verbose) {
            await logger.logToStderr('Retrieving group information...');
          }

          let group: Group;
          if (args.options.entraGroupId) {
            group = await entraGroup.getGroupById(args.options.entraGroupId);
          }
          else {
            group = await entraGroup.getGroupByDisplayName(args.options.entraGroupName!);
          }

          const entraSiteUser = await spo.ensureEntraGroup(args.options.webUrl, group);

          principalId = entraSiteUser.Id;

        }
        else {
          principalId = args.options.principalId!;
        }

        const serverRelativePath: string = urlUtil.getServerRelativePath(args.options.webUrl, fileURL);
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativePath)}')/ListItemAllFields/roleassignments/removeroleassignment(principalid='${principalId}')`,
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

    if (args.options.force) {
      await removeRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from file ${args.options.fileUrl || args.options.fileId} from site ${args.options.webUrl}?` });

      if (result) {
        await removeRoleAssignment();
      }
    }
  }

  private async getFileURL(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const file: FileProperties = await spo.getFileById(args.options.webUrl, args.options.fileId!, logger, this.verbose);
    return file.ServerRelativeUrl;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }
}

export default new SpoFileRoleAssignmentRemoveCommand();