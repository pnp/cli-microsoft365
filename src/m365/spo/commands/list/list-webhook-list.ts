import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        listId: (!(!args.options.listId)).toString(),
        listTitle: (!(!args.options.listTitle)).toString(),
        title: (!(!args.options.title)).toString()
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

        if (args.options.id && args.options.title) {
          return 'Specify id or title, but not both';
        }

        if (args.options.listId && args.options.listTitle) {
          return 'Specify listId or listTitle, but not both';
        }

        if (!args.options.id && !args.options.title) {
          if (!args.options.listId && !args.options.listTitle) {
            return 'Specify listId or listTitle, one is required';
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.id)}')/Subscriptions`;
    }
    else if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions`;
    }
    else if (args.options.listTitle) {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')/Subscriptions`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }>(requestOptions)
      .then((res: { value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }): void => {
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

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoListWebhookListCommand();