import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

const expirationDateTimeMaxDays = 180;
const maxExpirationDateTime: Date = new Date();
// 180 days from now is the maximum expiration date for a webhook
maxExpirationDateTime.setDate(maxExpirationDateTime.getDate() + expirationDateTimeMaxDays);

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  notificationUrl: string;
  expirationDateTime?: string;
  clientState?: string;
}

class SpoListWebhookAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_ADD;
  }

  public get description(): string {
    return 'Adds a new webhook to the specified list';
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
        expirationDateTime: (!(!args.options.expirationDateTime)).toString(),
        clientState: (!(!args.options.clientState)).toString()
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
        option: '-n, --notificationUrl <notificationUrl>'
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

        const parsedDateTime = Date.parse(args.options.expirationDateTime as string);
        if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
          return `Provide the date in one of the following formats:
          'YYYY-MM-DD'
          'YYYY-MM-DDThh:mm'
          'YYYY-MM-DDThh:mmZ'
          'YYYY-MM-DDThh:mm±hh:mm'`;
        }

        if (parsedDateTime < Date.now() || new Date(parsedDateTime) >= maxExpirationDateTime) {
          return `Provide an expiration date which is a date time in the future and within 6 months from now`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding webhook to list ${args.options.listId ? args.options.listId : args.options.listTitle} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/Subscriptions')`;
    }

    const requestBody: any = {};
    requestBody.resource = args.options.listId ? args.options.listId : args.options.listTitle;
    requestBody.notificationUrl = args.options.notificationUrl;
    // If no expiration date has been provided we will default to the
    // maximum expiration date of 180 days from now 
    requestBody.expirationDateTime = args.options.expirationDateTime
      ? new Date(args.options.expirationDateTime).toISOString()
      : maxExpirationDateTime.toISOString();
    if (args.options.clientState) {
      requestBody.clientState = args.options.clientState;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);

        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }
}

module.exports = new SpoListWebhookAddCommand();