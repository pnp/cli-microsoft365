import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
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
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional()
    .alias('i'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  listItemId: z.number(),
  principalId: z.number().optional(),
  upn: z.string().optional(),
  groupName: z.string().optional(),
  entraGroupId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `'${e.input}' is not a valid GUID for option entraGroupId.`
    })
    .optional(),
  entraGroupName: z.string().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRoleAssignmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEASSIGNMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes a role assignment from list item permissions';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length === 1, {
        error: 'Specify either listId, listTitle or listUrl.',
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      })
      .refine(opts => [opts.principalId, opts.upn, opts.groupName, opts.entraGroupId, opts.entraGroupName].filter(o => o !== undefined).length === 1, {
        error: 'Specify either principalId, upn, groupName, entraGroupId or entraGroupName.',
        params: {
          customCode: 'optionSet',
          options: ['principalId', 'upn', 'groupName', 'entraGroupId', 'entraGroupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removeRoleAssignment(logger, args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove role assignment from listitem ${args.options.listItemId} from list ${args.options.listId || args.options.listTitle} from site ${args.options.webUrl}?` });

      if (result) {
        await this.removeRoleAssignment(logger, args.options);
      }
    }
  }

  private async removeRoleAssignment(logger: Logger, options: Options): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Removing role assignment from listitem in site at ${options.webUrl}...`);
    }

    try {
      let requestUrl: string = `${options.webUrl}/_api/web/`;
      if (options.listId) {
        requestUrl += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')/`;
      }
      else if (options.listTitle) {
        requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
      }

      requestUrl += `items(${options.listItemId})/`;

      if (options.upn) {
        options.principalId = await this.getUserPrincipalId(options, logger);
      }
      else if (options.groupName) {
        options.principalId = await this.getGroupPrincipalId(options, logger);
      }
      else if (options.entraGroupId || options.entraGroupName) {
        if (this.verbose) {
          await logger.logToStderr('Retrieving group information...');
        }

        const group = options.entraGroupId
          ? await entraGroup.getGroupById(options.entraGroupId)
          : await entraGroup.getGroupByDisplayName(options.entraGroupName!);

        const siteUser = await spo.ensureEntraGroup(options.webUrl, group);
        options.principalId = siteUser.Id;
      }

      await this.removeRoleAssignmentWithRequestUrl(requestUrl, logger, options);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async removeRoleAssignmentWithRequestUrl(requestUrl: string, logger: Logger, options: Options): Promise<void> {
    const requestOptions: any = {
      url: `${requestUrl}roleassignments/removeroleassignment(principalid='${options.principalId}')`,
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

export default new SpoListItemRoleAssignmentRemoveCommand();