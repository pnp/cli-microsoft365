import { AxiosRequestConfig } from 'axios';
import * as chalk from 'chalk';
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
  id?: string;
  listTitle?: string;
  listId?: string;
  title?: string;
}

class SpoListWebhookListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_LIST;
  }

  public get description(): string {
    return 'Lists all webhooks for the specified list';
  }


  public defaultProperties(): string[] | undefined {
    return ['id', 'clientState', 'expirationDateTime', 'resource'];
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
        id: typeof args.options.id !== 'undefined',
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        title: typeof args.options.title !== 'undefined'
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
        option: '--id [id]'
      },
      {
        option: '--title [title]'
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

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
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
    this.optionSets.push(['id', 'title', 'listId', 'listTitle', 'listUrl']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.title && this.verbose) {
      logger.logToStderr(chalk.yellow(`Option 'title' is deprecated. Please use 'listTitle' instead`));
    }

    if (args.options.id && this.verbose) {
      logger.logToStderr(chalk.yellow(`Option 'id' is deprecated. Please use 'listId' instead`));
    }

    if (this.verbose) {
      const list: string = args.options.id ? formatting.encodeQueryParameter(args.options.id as string) : (args.options.listId ? formatting.encodeQueryParameter(args.options.listId as string) : (args.options.title ? formatting.encodeQueryParameter(args.options.title as string) : formatting.encodeQueryParameter(args.options.listTitle as string)));
      logger.logToStderr(`Retrieving webhook information for list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;
    if (args.options.id) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.id)}')/Subscriptions`;
    }
    else if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions`;
    }
    else if (args.options.title) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')/Subscriptions`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions`;
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
      const res = await request.get<{ value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }>(requestOptions);
      if (res.value && res.value.length > 0) {
        res.value.forEach(w => {
          w.clientState = w.clientState || '';
        });

        logger.log(res.value);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No webhooks found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListWebhookListCommand();