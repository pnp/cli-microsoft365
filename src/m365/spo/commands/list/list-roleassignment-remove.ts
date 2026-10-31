import { Group } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { entraGroup } from '../../../../utils/entraGroup.js';

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
      error: e => `'${e.input}' is not a valid number for option principalId.`
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
  force: z.boolean().alias('f').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from list permissions';
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeRoleAssignment = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Removing role assignment from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' of site ${args.options.webUrl}...`);
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

        let principalId: number | undefined = args.options.principalId;
        if (args.options.upn) {
          const user = await spo.ensureUser(args.options.webUrl, args.options.upn);
          principalId = user.Id;
        }
        else if (args.options.groupName) {
          const spGroup = await spo.getGroupByName(args.options.webUrl, args.options.groupName, logger, this.verbose);
          principalId = spGroup.Id;
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

        await this.removeRoleAssignment(requestUrl, principalId!);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeRoleAssignment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from the specified user of list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });

      if (result) {
        await removeRoleAssignment();
      }
    }
  }

  private async removeRoleAssignment(requestUrl: string, principalId: number): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}roleassignments/removeroleassignment(principalid='${principalId}')`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }
}

export default new SpoListRoleAssignmentRemoveCommand();