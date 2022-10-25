import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
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
          return `Specify camlQuery or fields but not both`;
        }

        if (args.options.camlQuery && args.options.pageSize) {
          return `Specify camlQuery or pageSize but not both`;
        }

        if (args.options.camlQuery && args.options.pageNumber) {
          return `Specify camlQuery or pageNumber but not both`;
        }

        if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
          return `pageSize must be numeric`;
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
      ['listId', 'listTitle']
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
    let requestUrl = `${args.options.webUrl}/_api/web`;

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

    let formDigestValue: string = '';

    const fieldsArray: string[] = args.options.fields ? args.options.fields.split(",")
      : (!args.options.output || args.options.output === "text") ? ["Title", "Id"] : [];

    const fieldsWithSlash: string[] = fieldsArray.filter(item => item.includes('/'));
    const fieldsToExpand: string[] = fieldsWithSlash.map(e => e.split('/')[0]);
    const expandFieldsArray: string[] = fieldsToExpand.filter((item, pos) => fieldsToExpand.indexOf(item) === pos);

    try {
      if (args.options.camlQuery) {
        if (this.debug) {
          logger.logToStderr(`getting request digest for query request`);
        }

        const res = await spo.getRequestDigest(args.options.webUrl);
        formDigestValue = res.FormDigestValue;
      }

      let res: any;
      if (args.options.pageNumber && Number(args.options.pageNumber) > 0) {
        const rowLimit: string = `$top=${Number(args.options.pageSize) * Number(args.options.pageNumber)}`;
        const filter: string = args.options.filter ? `$filter=${encodeURIComponent(args.options.filter)}` : ``;
        const fieldSelect: string = `?$select=Id&${rowLimit}&${filter}`;

        const requestOptions: AxiosRequestConfig = {
          url: `${requestUrl}/items${fieldSelect}`,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'X-RequestDigest': formDigestValue
          },
          responseType: 'json'
        };

        res = await request.get(requestOptions);
      }

      const skipTokenId = (res && res.value && res.value.length && res.value[res.value.length - 1]) ? res.value[res.value.length - 1].Id : 0;
      const skipToken: string = (args.options.pageNumber && Number(args.options.pageNumber) > 0 && skipTokenId > 0) ? `$skiptoken=Paged=TRUE%26p_ID=${res.value[res.value.length - 1].Id}` : ``;
      const rowLimit: string = args.options.pageSize ? `$top=${args.options.pageSize}` : ``;
      const filter: string = args.options.filter ? `$filter=${encodeURIComponent(args.options.filter)}` : ``;
      const fieldExpand: string = expandFieldsArray.length > 0 ? `&$expand=${expandFieldsArray.join(",")}` : ``;
      const fieldSelect: string = fieldsArray.length > 0 ?
        `?$select=${encodeURIComponent(fieldsArray.join(","))}${fieldExpand}&${rowLimit}&${skipToken}&${filter}` :
        `?${rowLimit}&${skipToken}&${filter}`;
      const requestBody: any = args.options.camlQuery ?
        {
          "query": {
            "ViewXml": args.options.camlQuery
          }
        }
        : ``;

      const requestOptions: AxiosRequestConfig = {
        url: `${requestUrl}/${args.options.camlQuery ? `GetItems` : `items${fieldSelect}`}`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'X-RequestDigest': formDigestValue
        },
        responseType: 'json',
        data: requestBody
      };

      const listItemInstances = args.options.camlQuery ? await request.post<ListItemInstanceCollection>(requestOptions) : await request.get<ListItemInstanceCollection>(requestOptions);
      listItemInstances.value.forEach(v => delete v['ID']);
      logger.log(listItemInstances.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListItemListCommand();