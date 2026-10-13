import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listTitle: z.string().optional().alias('t'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('i'),
  listUrl: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFieldListCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_LIST;
  }

  public get description(): string {
    return 'Retrieves columns for the specified list or site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'InternalName', 'Hidden'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(o => o !== undefined).length <= 1, {
        error: 'Specify either list id or title or list url'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let listUrl: string = '';

    if (args.options.listId) {
      listUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    try {
      const res = await odata.getAllItems<any>(`${args.options.webUrl}/_api/web/${listUrl}fields`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFieldListCommand();