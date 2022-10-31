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
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
  clientState?: string;
  id: string;
}

class SpoListWebhookSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_SET;
  }

  public get description(): string {
    return 'Updates the specified webhook';
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
        notificationUrl: typeof args.options.notificationUrl !== 'undefined',
        expirationDateTime: typeof args.options.expirationDateTime !== 'undefined',
        clientState: typeof args.options.clientState !== 'undefined'
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
        option: '-i, --id <id>'
      },
      {
        option: '-n, --notificationUrl [notificationUrl]'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]'
      },
      {
        option: '-c, --clientState [clientState]'
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

        if (!args.options.notificationUrl && !args.options.expirationDateTime && !args.options.clientState) {
          return 'Specify notificationUrl, expirationDateTime, clientState or multiple, at least one is required';
        }

        const parsedDateTime = Date.parse(args.options.expirationDateTime as string);
        if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
          return `${args.options.expirationDateTime} is not a valid date format. Provide the date in one of the following formats:
      ${chalk.grey('YYYY-MM-DD')}
      ${chalk.grey('YYYY-MM-DDThh:mm')}
      ${chalk.grey('YYYY-MM-DDThh:mmZ')}
      ${chalk.grey('YYYY-MM-DDThh:mm±hh:mm')}`;
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
      logger.logToStderr(`Updating webhook ${args.options.id} belonging to list ${args.options.listId || args.options.listTitle || args.options.listUrl} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }

    const requestBody: any = {};
    if (args.options.notificationUrl) {
      requestBody.notificationUrl = args.options.notificationUrl;
    }
    if (args.options.expirationDateTime) {
      requestBody.expirationDateTime = args.options.expirationDateTime;
    }
    if (args.options.clientState) {
      requestBody.clientState = args.options.clientState;
    }

    const requestOptions: AxiosRequestConfig = {
      url: requestUrl,
      method: 'PATCH',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
      // REST patch call doesn't return anything
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListWebhookSetCommand();