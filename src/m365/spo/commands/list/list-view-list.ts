import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
}

class SpoListViewListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_LIST;
  }

  public get description(): string {
    return 'Lists views configured on the specified list';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'DefaultView', 'Hidden', 'BaseViewId'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined'
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

        if (args.options.listId) {
          if (!validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/views`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/views`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/views`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      logger.log(res.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListViewListCommand(); 