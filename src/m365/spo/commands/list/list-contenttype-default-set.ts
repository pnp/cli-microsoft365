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
  listId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `${e.input} is not a valid GUID`
  }).optional().alias('l'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  contentTypeId: z.string().alias('c')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface ContentTypes {
  value: {
    Id: StringValue;
  }[];
}

interface StringValue {
  StringValue: string;
}

class SpoListContentTypeDefaultSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_CONTENTTYPE_DEFAULT_SET;
  }

  public get description(): string {
    return 'Sets the default content type for a list';
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

    let baseUrl: string = `${args.options.webUrl}/_api/web/`;
    if (args.options.listId) {
      baseUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      baseUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      baseUrl += `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    if (this.verbose) {
      await logger.logToStderr('Retrieving content type order...');
    }

    try {
      const contentTypeOrder: StringValue[] = await this.getContentTypeOrder(baseUrl, logger);

      // see if the specified content type is among the registered content types
      // if it is, it means it's visible
      const contentTypeIndex: number = contentTypeOrder.findIndex(ct => ct.StringValue.toUpperCase() === args.options.contentTypeId.toUpperCase());

      if (contentTypeIndex > -1) {
        if (this.debug) {
          await logger.logToStderr(`Content type ${args.options.contentTypeId} is visible in the list`);
        }
        // content type is in the list and is visible in the menu

        if (contentTypeIndex === 0) {
          if (this.verbose) {
            await logger.logToStderr(`Content type ${args.options.contentTypeId} is already set as default`);
          }
        }
        else {
          if (this.verbose) {
            await logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
          }

          // remove content type from the order array so that we can put it at
          // the beginning to make it default content type          
          contentTypeOrder.splice(contentTypeIndex, 1);
          contentTypeOrder.unshift({
            StringValue: args.options.contentTypeId
          });

          await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
        }
      }
      else {
        if (this.debug) {
          await logger.logToStderr(`Content type ${args.options.contentTypeId} is not visible in the list`);
        }

        if (this.verbose) {
          await logger.logToStderr('Retrieving list content types...');
        }

        const contentTypes: string[] = await this.getListContentTypes(baseUrl);
        if (!contentTypes.find(ct => ct.toUpperCase() === args.options.contentTypeId.toUpperCase())) {
          throw `Content type ${args.options.contentTypeId} missing in the list. Add the content type to the list first and try again.`;
        }

        if (this.verbose) {
          await logger.logToStderr(`Setting content type ${args.options.contentTypeId} as default...`);
        }

        contentTypeOrder.unshift({
          StringValue: args.options.contentTypeId
        });

        await this.updateContentTypeOrder(baseUrl, contentTypeOrder);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getContentTypeOrder(baseUrl: string, logger: Logger): Promise<StringValue[]> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/RootFolder?$select=ContentTypeOrder,UniqueContentTypeOrder`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ ContentTypeOrder: StringValue[]; UniqueContentTypeOrder: StringValue[] | null }>(requestOptions);
    let uniqueContentTypeOrder = response.ContentTypeOrder;
    if (response.UniqueContentTypeOrder !== null) {
      if (this.debug) {
        await logger.logToStderr('Using unique content type order');
      }
      uniqueContentTypeOrder = response.UniqueContentTypeOrder as StringValue[];
    }
    else {
      if (this.debug) {
        await logger.logToStderr('Unique content type order not defined. Using content type order');
      }
    }
    return uniqueContentTypeOrder;
  }

  private async updateContentTypeOrder(baseUrl: string, contentTypeOrder: StringValue[]): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/RootFolder`,
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'x-http-method': 'MERGE'
      },
      data: {
        UniqueContentTypeOrder: contentTypeOrder
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  private async getListContentTypes(baseUrl: string): Promise<string[]> {
    const requestOptions: CliRequestOptions = {
      url: `${baseUrl}/ContentTypes?$select=Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };
    const response = await request.get<ContentTypes>(requestOptions);
    return response.value.map(ct => ct.Id.StringValue);
  }
}

export default new SpoListContentTypeDefaultSetCommand();