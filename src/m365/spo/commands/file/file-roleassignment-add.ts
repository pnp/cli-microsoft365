import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { RoleDefinition } from '../roledefinition/RoleDefinition.js';
import { FileProperties } from './FileProperties.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  fileUrl: z.string().optional(),
  fileId: z.uuid().optional().alias('i'),
  principalId: z.number().optional(),
  upn: z.string().optional(),
  groupName: z.string().optional(),
  entraGroupId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID for option entraGroupId.`
  }).optional(),
  entraGroupName: z.string().optional(),
  roleDefinitionId: z.number().optional(),
  roleDefinitionName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFileRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to the specified file.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.fileId, opts.fileUrl].filter(o => o !== undefined).length === 1, {
        error: `Specify 'fileId' or 'fileUrl', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['fileId', 'fileUrl']
        }
      })
      .refine(opts => [opts.principalId, opts.upn, opts.groupName, opts.entraGroupId, opts.entraGroupName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'principalId', 'upn', 'groupName', 'entraGroupId', or 'entraGroupName'.`,
        params: {
          customCode: 'optionSet',
          options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName']
        }
      })
      .refine(opts => [opts.roleDefinitionId, opts.roleDefinitionName].filter(o => o !== undefined).length === 1, {
        error: `Specify either 'roleDefinitionId' or 'roleDefinitionName'.`,
        params: {
          customCode: 'optionSet',
          options: ['roleDefinitionId', 'roleDefinitionName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding role assignment to file in site at ${args.options.webUrl}...`);
    }

    try {
      const fileUrl: string = await this.getFileURL(args, logger);
      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      if (args.options.upn) {
        const upnPrincipalId = await this.getUserPrincipalId(args.options, logger);
        await this.addRoleAssignment(fileUrl, args.options.webUrl, upnPrincipalId, roleDefinitionId);
      }
      else if (args.options.groupName) {
        const groupPrincipalId = await this.getGroupPrincipalId(args.options, logger);
        await this.addRoleAssignment(fileUrl, args.options.webUrl, groupPrincipalId, roleDefinitionId);
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
        await this.addRoleAssignment(fileUrl, args.options.webUrl, entraSiteUser.Id, roleDefinitionId);
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
      url: `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(fileUrl)}')/ListItemAllFields/roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
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
      return options.roleDefinitionId!;
    }

    const roleDefinition: RoleDefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName, logger, this.verbose);
    return roleDefinition.Id;
  }

  private async getGroupPrincipalId(options: Options, logger: Logger): Promise<number> {
    const group = await spo.getGroupByName(options.webUrl, options.groupName!, logger, this.verbose);
    return group.Id;
  }

  private async getUserPrincipalId(options: Options, logger: Logger): Promise<number> {
    const user = await spo.getUserByEmail(options.webUrl, options.upn!, logger, this.verbose);
    return user.Id;
  }

  private async getFileURL(args: CommandArgs, logger: Logger): Promise<string> {
    if (args.options.fileUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.fileUrl);
    }

    const file: FileProperties = await spo.getFileById(args.options.webUrl, args.options.fileId!, logger, this.verbose);
    return file.ServerRelativeUrl;
  }
}

export default new SpoFileRoleAssignmentAddCommand();
