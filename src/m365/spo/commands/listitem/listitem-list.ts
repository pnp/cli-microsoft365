import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstance } from './ListItemInstance';
import { ListItemInstanceCollection } from './ListItemInstanceCollection';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  fields?: string;
  filter?: string;
  pageNumber?: string;
  pageSize?: string;
  camlQuery?: string;
  webUrl: string;
}

class SpoListItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Gets a list of items from the specified list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        fields: typeof args.options.fields !== 'undefined',
        filter: typeof args.options.filter !== 'undefined',
        pageNumber: typeof args.options.pageNumber !== 'undefined',
        pageSize: typeof args.options.pageSize !== 'undefined',
        camlQuery: typeof args.options.camlQuery !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-s, --pageSize [pageSize]'
      },
      {
        option: '-n, --pageNumber [pageNumber]'
      },
      {
        option: '-q, --camlQuery [camlQuery]'
      },
      {
        option: '-f, --fields [fields]'
      },
      {
        option: '-l, --filter [filter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.camlQuery && args.options.fields) {
          return `You cannot use the fields-option when specifying a camlQuery`;
        }

        if (args.options.camlQuery && args.options.pageSize) {
          return `You cannot use the pageSize-option when specifying a camlQuery`;
        }

        if (args.options.camlQuery && args.options.pageNumber) {
          return `You cannot use the pageNumber-option when specifying a camlQuery`;
        }

        if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
          return `pageSize ${args.options.pageSize} must be numeric`;
        }

        if (args.options.pageNumber && !args.options.pageSize) {
          return `pageSize must be specified if pageNumber is specified`;
        }

        if (args.options.pageNumber && isNaN(Number(args.options.pageNumber))) {
          return `pageNumber must be numeric`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'camlQuery',
      'pageSize',
      'pageNumber',
      'fields',
      'filter'
    );
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
      logger.log(listItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getItems(logger: Logger, options: Options, listApiUrl: string): Promise<ListItemInstance[]> {
    if (this.verbose) {
      logger.logToStderr(`Getting list items`);
    }

    const queryParams = [];
    const fieldsArray: string[] = options.fields ? options.fields.split(",")
      : (!options.output || Cli.shouldTrimOutput(options.output)) ? ["Title", "Id"] : [];
    const expandFieldsArray: string[] = this.getExpandFieldsArray(fieldsArray);
    const skipTokenId = await this.getSkipTokenId(logger, options, listApiUrl);

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
    const formDigestValue = await this.getFormDigestValue(logger, options.webUrl);

    if (this.verbose) {
      logger.logToStderr(`Getting list items using CAML query`);
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

  private async getFormDigestValue(logger: Logger, webUrl: string): Promise<string> {
    if (this.verbose) {
      logger.logToStderr(`Getting request digest for query request`);
    }

    const response = await spo.getRequestDigest(webUrl);
    return response.FormDigestValue;
  }

  private async getSkipTokenId(logger: Logger, options: Options, listApiUrl: string): Promise<number | undefined> {
    if (!options.pageNumber || Number(options.pageNumber) === 0) {
      return undefined;
    }

    if (this.verbose) {
      logger.logToStderr(`Getting skipToken Id for page ${options.pageNumber}`);
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

module.exports = new SpoListItemListCommand();