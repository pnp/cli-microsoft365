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
  webUrl: z.string().refine(url => validation.isValidSharePointUrl(url) === true, {
    error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
  }).alias('u'),
  listTitle: z.string().optional().alias('l'),
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional(),
  listUrl: z.string().optional(),
  id: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('i'),
  title: z.string().optional().alias('t'),
  internalName: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoFieldGetCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_GET;
  }

  public get description(): string {
    return 'Retrieves information about the specified list- or site column';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.id, opts.title, opts.internalName].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: id, title, internalName.',
        params: {
          customCode: 'optionSet',
          options: ['id', 'title', 'internalName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let listRestUrl: string = '';

    if (args.options.listId) {
      listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/`;
    }

    let fieldRestUrl: string;
    if (args.options.id) {
      fieldRestUrl = `/getbyid('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else {
      fieldRestUrl = `/getbyinternalnameortitle('${formatting.encodeQueryParameter((args.options.title || args.options.internalName) as string)}')`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/${listRestUrl}fields${fieldRestUrl}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoFieldGetCommand();