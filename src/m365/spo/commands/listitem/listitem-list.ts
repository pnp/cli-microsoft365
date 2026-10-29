import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListItemInstance } from './ListItemInstance.js';
import { ListItemInstanceCollection } from './ListItemInstanceCollection.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(webUrl => validation.isValidSharePointUrl(webUrl) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  listId: z.string()
    .refine(listId => validation.isValidGuid(listId), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional()
    .alias('i'),
  listTitle: z.string().optional().alias('t'),
  listUrl: z.string().optional(),
  pageSize: z.number().int().positive().optional().alias('s'),
  pageNumber: z.number().int().nonnegative().optional().alias('n'),
  camlQuery: z.string().optional().alias('q'),
  fields: z.string().optional(),
  filter: z.string().optional().alias('l')
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Gets a list of items from the specified list';
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
      .refine(opts => !opts.camlQuery || !opts.fields, {
        error: 'You cannot use the fields-option when specifying a camlQuery'
      })
      .refine(opts => !opts.camlQuery || !opts.pageSize, {
        error: 'You cannot use the pageSize-option when specifying a camlQuery'
      })
      .refine(opts => !opts.camlQuery || !opts.pageNumber, {
        error: 'You cannot use the pageNumber-option when specifying a camlQuery'
      })
      .refine(opts => !opts.pageNumber || opts.pageSize !== undefined, {
        error: 'pageSize must be specified if pageNumber is specified'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let listApiUrl = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      listApiUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      listApiUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      listApiUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
    }

    try {
      const listItems = args.options.camlQuery
        ? await this.getItemsUsingCAMLQuery(logger, args.options, listApiUrl)
        : await this.getItems(logger, args.options, listApiUrl);

      listItems.forEach(v => delete v['ID']);
      await logger.log(listItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getItems(logger: Logger, options: Options, listApiUrl: string): Promise<ListItemInstance[]> {
    if (this.verbose) {
      await logger.logToStderr(`Getting list items`);
    }

    const queryParams = [];
    const fieldsArray: string[] = options.fields ? options.fields.split(",")
      : (!options.output || cli.shouldTrimOutput(options.output)) ? ["Title", "Id"] : [];
    const expandFieldsArray: string[] = this.getExpandFieldsArray(fieldsArray);
    const skipTokenId = await this.getLastItemIdForPage(logger, options, listApiUrl);

    queryParams.push(`$top=${options.pageSize || 5000}`);

    if (options.filter) {
      queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
    }

    if (expandFieldsArray.length > 0) {
      queryParams.push(`$expand=${expandFieldsArray.join(",")}`);
    }

    if (fieldsArray.length > 0) {
      queryParams.push(`$select=${formatting.encodeQueryParameter(fieldsArray.join(","))}`);
    }

    if (skipTokenId !== undefined) {
      queryParams.push(`$skiptoken=Paged=TRUE%26p_ID=${skipTokenId}`);
    }

    // If skiptoken is not found, then we are past the last page
    if (options.pageNumber && Number(options.pageNumber) > 0 && skipTokenId === undefined) {
      return [];
    }

    if (!options.pageSize) {
      return await odata.getAllItems<ListItemInstance>(`${listApiUrl}/items?${queryParams.join('&')}`);
    }
    else {
      const requestOptions: CliRequestOptions = {
        url: `${listApiUrl}/items?${queryParams.join('&')}`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const listItemCollection = await request.get<ListItemInstanceCollection>(requestOptions);
      return listItemCollection.value;
    }
  }

  private async getItemsUsingCAMLQuery(logger: Logger, options: Options, listApiUrl: string): Promise<ListItemInstance[]> {
    const formDigestValue = (await spo.getRequestDigest(options.webUrl)).FormDigestValue;

    if (this.verbose) {
      await logger.logToStderr(`Getting list items using CAML query`);
    }

    const items: ListItemInstance[] = [];
    let skipTokenId: number | undefined = undefined;

    do {
      const requestBody: any = {
        "query": {
          "ViewXml": options.camlQuery,
          "AllowIncrementalResults": true
        }
      };

      if (skipTokenId !== undefined) {
        requestBody.query.ListItemCollectionPosition = {
          "PagingInfo": `Paged=TRUE&p_ID=${skipTokenId}`
        };
      }

      const requestOptions: CliRequestOptions = {
        url: `${listApiUrl}/GetItems`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'X-RequestDigest': formDigestValue
        },
        responseType: 'json',
        data: requestBody
      };

      const listItemInstances = await request.post<ListItemInstanceCollection>(requestOptions);
      skipTokenId = listItemInstances.value.length > 0 ? listItemInstances.value[listItemInstances.value.length - 1].Id : undefined;
      items.push(...listItemInstances.value);
    }
    while (skipTokenId !== undefined);

    return items;
  }

  private getExpandFieldsArray(fieldsArray: string[]): string[] {
    const fieldsWithSlash: string[] = fieldsArray.filter(item => item.includes('/'));
    const fieldsToExpand: string[] = fieldsWithSlash.map(e => e.split('/')[0]);
    const expandFieldsArray: string[] = fieldsToExpand.filter((item, pos) => fieldsToExpand.indexOf(item) === pos);
    return expandFieldsArray;
  }

  private async getLastItemIdForPage(logger: Logger, options: Options, listApiUrl: string): Promise<number | undefined> {
    if (!options.pageNumber || Number(options.pageNumber) === 0) {
      return undefined;
    }

    if (this.verbose) {
      await logger.logToStderr(`Getting skipToken Id for page ${options.pageNumber}`);
    }

    const rowLimit: string = `$top=${Number(options.pageSize) * Number(options.pageNumber)}`;
    const filter: string = options.filter ? `$filter=${encodeURIComponent(options.filter)}` : ``;

    const requestOptions: CliRequestOptions = {
      url: `${listApiUrl}/items?$select=Id&${rowLimit}&${filter}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ Id: number }] }>(requestOptions);
    return response.value[response.value.length - 1]?.Id;
  }
}

export default new SpoListItemListCommand();