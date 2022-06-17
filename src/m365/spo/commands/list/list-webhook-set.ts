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
  listId?: string;
  listTitle?: string;
  notificationUrl?: string;
  expirationDateTime?: string;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: (!(!args.options.listId)).toString(),
        listTitle: (!(!args.options.listTitle)).toString(),
        notificationUrl: (!(!args.options.notificationUrl)).toString(),
        expirationDateTime: (!(!args.options.expirationDateTime)).toString()
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
        option: '-i, --id <id>'
      },
      {
        option: '-n, --notificationUrl [notificationUrl]'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]'
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

        if (args.options.listId && args.options.listTitle) {
          return 'Specify listId or listTitle, but not both';
        }

        if (!args.options.listId && !args.options.listTitle) {
          return 'Specify listId or listTitle, one is required';
        }

        if (!args.options.notificationUrl && !args.options.expirationDateTime) {
          return 'Specify notificationUrl, expirationDateTime or both, at least one is required';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Updating webhook ${args.options.id} belonging to list ${args.options.listId ? args.options.listId : args.options.listTitle} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }

    const requestBody: any = {};
    if (args.options.notificationUrl) {
      requestBody.notificationUrl = args.options.notificationUrl;
    }
    if (args.options.expirationDateTime) {
      requestBody.expirationDateTime = args.options.expirationDateTime;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'PATCH',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    request
      .patch(requestOptions)
      .then((): void => {
        // REST patch call doesn't return anything
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }
}

module.exports = new SpoListWebhookSetCommand();