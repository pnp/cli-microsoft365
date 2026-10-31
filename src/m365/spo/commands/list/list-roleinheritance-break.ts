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
      error: e => `${e.input} is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .alias('i')
    .optional(),
  listTitle: z.string().alias('t').optional(),
  listUrl: z.string().optional(),
  clearExistingPermissions: z.boolean().alias('c').optional(),
  force: z.boolean().alias('f').optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks role inheritance on list or library';
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
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Breaking role inheritance of list in site at ${args.options.webUrl}...`);
    }

    const breakListRoleInheritance = async (): Promise<void> => {
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

        let keepExistingPermissions: boolean = true;
        if (args.options.clearExistingPermissions) {
          keepExistingPermissions = !args.options.clearExistingPermissions;
        }

        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}breakroleinheritance(${keepExistingPermissions})`,
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
    };

    if (args.options.force) {
      await breakListRoleInheritance();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to break the role inheritance of ${args.options.listId ?? args.options.listTitle}?` });

      if (result) {
        await breakListRoleInheritance();
      }
    }
  }
}

export default new SpoListRoleInheritanceBreakCommand();