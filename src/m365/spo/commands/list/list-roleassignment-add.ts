import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option listId.`
    })
    .alias('i')
    .optional(),
  listTitle: z.string().alias('t').optional(),
  listUrl: z.string().optional(),
  principalId: z.number()
    .refine(id => validation.isValidPositiveInteger(id), {
      error: e => `Specified principalId '${e.input}' is not a valid number.`
    })
    .optional(),
  upn: z.string()
    .refine(upn => validation.isValidUserPrincipalName(upn), {
      error: e => `'${e.input}' is not a valid user principal name for option upn.`
    })
    .optional(),
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

class SpoListRoleAssignmentAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEASSIGNMENT_ADD;
  }

  public get description(): string {
    return 'Adds a role assignment to list permissions';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: listId, listTitle, listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      })
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
      await logger.logToStderr(`Adding role assignment to list in site at ${args.options.webUrl}...`);
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

      const roleDefinitionId = await this.getRoleDefinitionId(args.options, logger);
      let principalId: number | undefined = args.options.principalId;
      if (args.options.upn) {
        const user = await spo.ensureUser(args.options.webUrl, args.options.upn);
        principalId = user.Id;
      }
      else if (args.options.groupName) {
        const group = await spo.getGroupByName(args.options.webUrl, args.options.groupName, logger, this.verbose);
        principalId = group.Id;
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
      url: `${requestUrl}roleassignments/addroleassignment(principalid='${principalId}',roledefid='${roleDefinitionId}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getRoleDefinitionId(options: Options, logger: Logger): Promise<number> {
    if (options.roleDefinitionId) {
      return options.roleDefinitionId;
    }

    const roleDefinition = await spo.getRoleDefinitionByName(options.webUrl, options.roleDefinitionName!, logger, this.verbose);
    return roleDefinition.Id;
  }
}

export default new SpoListRoleAssignmentAddCommand();