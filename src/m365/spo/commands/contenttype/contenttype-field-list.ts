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
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  contentTypeId: z.string().optional().alias('i'),
  contentTypeName: z.string().optional().alias('n'),
  listTitle: z.string().optional().alias('l'),
  listId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID for option 'listId'.`
    }).optional(),
  listUrl: z.string().optional(),
  properties: z.string().optional().alias('p')
});

export type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoContentTypeFieldListCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPE_FIELD_LIST;
  }

  public get description(): string {
    return 'Lists fields for a given site or list content type';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'InternalName', 'Hidden'];
  }


  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(opts => [opts.contentTypeId, opts.contentTypeName].filter(x => x !== undefined).length === 1, {
        message: `Specify either 'contentTypeId' or 'contentTypeName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['contentTypeId', 'contentTypeName']
        }
      })
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(x => x !== undefined).length <= 1, {
        message: `Specify either 'listId', 'listTitle' or 'listUrl'.`,
        params: {
          customCode: 'optionSet',
          options: ['listId', 'listTitle', 'listUrl']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving fields for content type '${args.options.contentTypeId || args.options.contentTypeName}' in site ${args.options.webUrl}...`);
      }

      let requestUrl: string = `${args.options.webUrl}/_api/web`;
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

      requestUrl += '/contentTypes';

      const contentTypeId = await this.getContentTypeId(requestUrl, logger, args.options.contentTypeId, args.options.contentTypeName);
      requestUrl += `('${formatting.encodeQueryParameter(contentTypeId)}')/fields`;

      if (args.options.properties) {
        requestUrl += `?$select=${args.options.properties}`;
      }

      const res = await odata.getAllItems(requestUrl);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeId(requestUrl: string, logger: Logger, contentTypeId?: string, contentTypeName?: string): Promise<string> {
    if (contentTypeId) {
      return contentTypeId;
    }

    if (this.verbose) {
      await logger.logToStderr(`Retrieving content type id for content type '${contentTypeName}'.`);
    }

    const res: { StringId: string }[] = await odata.getAllItems(`${requestUrl}?$filter=Name eq '${formatting.encodeQueryParameter(contentTypeName!)}'&$select=StringId`);

    if (res.length === 0) {
      throw `Content type with name ${contentTypeName} not found.`;
    }

    return res[0].StringId;
  }
}

export default new SpoContentTypeFieldListCommand();