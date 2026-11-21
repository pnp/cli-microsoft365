import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { RoleAssignment } from '../roledefinition/RoleDefinition.js';

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
  listItemId: z.number().int().positive()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ListItemRoleAssignment extends RoleAssignment {
  PrincipalId: number;
}

class SpoListItemRoleAssignmentListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists all role assignments of a list item';
  }

  public defaultProperties(): string[] | undefined {
    return ['PrincipalId', 'Member', 'RoleDefinitionBindings'];
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving role assignments of list item with ID ${args.options.listItemId}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}/items(${args.options.listItemId})/RoleAssignments?$expand=Member,RoleDefinitionBindings`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const response = await request.get<{ value: ListItemRoleAssignment[] }>(requestOptions);
      response.value.forEach(r => {
        r.RoleDefinitionBindings = formatting.setFriendlyPermissions(r.RoleDefinitionBindings);
      });

      if (args.options.output && args.options.output !== 'json') {
        await logger.log(response.value.map(r => ({
          PrincipalId: r.PrincipalId,
          Member: r.Member.Title,
          RoleDefinitionBindings: r.RoleDefinitionBindings.map(d => d.Name).join(', ')
        })));
      }
      else {
        await logger.log(response.value);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemRoleAssignmentListCommand();
