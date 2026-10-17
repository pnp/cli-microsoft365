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
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  listTitle: z.string().optional(),
  listUrl: z.string().optional(),
  viewId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  viewTitle: z.string().optional(),
  id: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  title: z.string().optional(),
  position: z.number()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListViewFieldSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_SET;
  }

  public get description(): string {
    return 'Updates existing column in an existing view (eg. move to a specific position).';
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
      .refine(opts => [opts.viewId, opts.viewTitle].filter(o => o !== undefined).length === 1, {
        error: 'Specify one of the following options: viewId, viewTitle.',
        params: {
          customCode: 'optionSet',
          options: ['viewId', 'viewTitle']
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
    let listSelector: string = '';
    if (args.options.listId) {
      listSelector = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      listSelector = `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      listSelector = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;

    if (this.verbose) {
      await logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
    }

    try {
      const field = await this.getField(args.options, listSelector);

      if (this.verbose) {
        await logger.logToStderr(`Moving the field ${args.options.id || args.options.title} in view ${args.options.viewId || args.options.viewTitle} to position ${args.options.position}...`);
      }

      const moveRequestUrl: string = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;
      const moveRequestOptions: CliRequestOptions = {
        url: moveRequestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          field: field.InternalName,
          index: args.options.position
        },
        responseType: 'json'
      };

      await request.post(moveRequestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getField(options: Options, listSelector: string): Promise<{ InternalName: string; }> {
    const fieldSelector: string = options.id ? `/getbyid('${formatting.encodeQueryParameter(options.id)}')` : `/getbyinternalnameortitle('${formatting.encodeQueryParameter(options.title as string)}')`;
    const getRequestUrl: string = `${options.webUrl}/_api/web/${listSelector}/fields${fieldSelector}`;

    const requestOptions: CliRequestOptions = {
      url: getRequestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }
}

export default new SpoListViewFieldSetCommand();