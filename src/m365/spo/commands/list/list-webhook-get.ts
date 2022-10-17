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
  id: string;
}

class SpoListWebhookGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_GET;
  }

  public get description(): string {
    return 'Gets information about the specific webhook';
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
        listUrl: typeof args.options.listUrl !== 'undefined',
        id: typeof args.options.id !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-i, --id [id]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

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
    this.optionSets.push(['listId', 'listTitle', 'listUrl']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      const list: string = (args.options.listId ? args.options.listId : args.options.listTitle) as string;
      logger.logToStderr(`Retrieving information for webhook ${args.options.id} belonging to list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }

    const requestOptions: AxiosRequestConfig = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      if (this.verbose) {
        logger.logToStderr('Specified webhook not found');
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListWebhookGetCommand();