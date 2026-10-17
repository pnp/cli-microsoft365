import { z } from 'zod';
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
      error: e => `${e.input} in option listId is not a valid GUID`
    })
    .optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} in option id is not a valid GUID`
    })
    .optional(),
  title: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListViewGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_GET;
  }

  public get description(): string {
    return 'Gets information about specific list view';
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
      .refine(opts => [opts.id, opts.title].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: id, title.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const baseRestUrl: string = `${args.options.webUrl}/_api/web`;
    let listRestUrl: string = '';

    if (args.options.listId) {
      listRestUrl = `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const viewRestUrl: string = `/views/${(args.options.id ? `getById('${formatting.encodeQueryParameter(args.options.id)}')` : `getByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`)}`;

    const requestOptions: CliRequestOptions = {
      url: `${baseRestUrl}${listRestUrl}${viewRestUrl}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const result = await request.get<any>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListViewGetCommand();