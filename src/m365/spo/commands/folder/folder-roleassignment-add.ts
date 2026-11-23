import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { spo } from '../../../../utils/spo.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  folderUrl: z.string(),
  principalId: z.number()
    .refine(id => validation.isValidPositiveInteger(id), {
      error: e => `Specified principalId '${e.input}' is not a valid number.`
    })
    .optional(),
  upn: z.string().optional(),
  groupName: z.string().optional(),
  entraGroupId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option entraGroupId.`
    })
    .optional(),
  entraGroupName: z.string().optional(),
  roleDefinitionId: z.number()
    .refine(id => validation.isValidPositiveInteger(id), {
      error: e => `Specified roleDefinitionId '${e.input}' is not a valid number.`
    })
    .optional(),
  roleDefinitionName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFolderRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to the specified folder.';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.principalId, opts.upn, opts.groupName, opts.entraGroupId, opts.entraGroupName].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: principalId, upn, groupName, entraGroupId, entraGroupName.',
        params: {
          customCode: 'optionSet',
          options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName']
        }
      })
      .refine(opts => [opts.roleDefinitionId, opts.roleDefinitionName].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: roleDefinitionId, roleDefinitionName.',
        params: {
          customCode: 'optionSet',
          options: ['roleDefinitionId', 'roleDefinitionName']
        }
      });
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