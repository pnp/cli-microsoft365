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
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  id: z.string().alias('i'),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListContentTypeRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_REMOVE;
  }

  public get description(): string {
    return 'Removes content type from list';
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
    const removeContentTypeFromList = async (): Promise<void> => {
      if (this.verbose) {
        const list: string = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl) as string;
        await logger.logToStderr(`Removing content type ${args.options.id} from list ${list} in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web/`;

      if (args.options.listId) {
        requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${requestUrl}/ContentTypes('${formatting.encodeQueryParameter(args.options.id)}')`,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeContentTypeFromList();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the content type ${args.options.id} from the list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}?` });

      if (result) {
        await removeContentTypeFromList();
      }
    }
  }
}

export default new SpoListContentTypeRemoveCommand();