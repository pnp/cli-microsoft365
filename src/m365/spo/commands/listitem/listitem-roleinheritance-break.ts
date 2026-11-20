import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listItemId: z.number(),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional()
    .alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  clearExistingPermissions: z.boolean().optional().alias('c'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks inheritance of list item';
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
      await logger.logToStderr(`Breaking role inheritance of list item in site at ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.breakListItemRoleInheritance(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?` });

      if (result) {
        await this.breakListItemRoleInheritance(args.options);
      }
    }
  }

  private async breakListItemRoleInheritance(options: Options): Promise<void> {
    try {
      let requestUrl: string = `${options.webUrl}/_api/web`;

      if (options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
      }
      else if (options.listTitle) {
        requestUrl += `/lists/getbytitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      let keepExistingPermissions: boolean = true;
      if (options.clearExistingPermissions) {
        keepExistingPermissions = !options.clearExistingPermissions;
      }

      const requestOptions: CliRequestOptions = {
        url: `${requestUrl}/items(${options.listItemId})/breakroleinheritance(${keepExistingPermissions})`,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemRoleInheritanceBreakCommand();