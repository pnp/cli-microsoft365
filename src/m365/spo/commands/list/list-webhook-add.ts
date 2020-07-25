import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { CommandInstance } from '../../../../cli';

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.expirationDateTime = (!(!args.options.expirationDateTime)).toString();
    telemetryProps.clientState = (!(!args.options.clientState)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Adding webhook to list ${args.options.listId ? encodeURIComponent(args.options.listId) : encodeURIComponent(args.options.listTitle as string)} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions')`;
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
      body: requestBody,
      json: true
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to add the webhook to is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list to which the webhook which should be added. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list to which the webhook which should be added. Specify either listId or listTitle but not both'
      },
      {
        option: '-n, --notificationUrl <notificationUrl>',
        description: 'The notification url'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]',
        description: 'The expiration date. Will be set to max (6 months from today) if not provided.'
      },
      {
        option: '-c, --clientState [clientState]',
        description: 'A client state information that will be passed through notifications.'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId) {
        if (!Utils.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      const parsedDateTime = Date.parse(args.options.expirationDateTime as string)
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
    };
  }
}

module.exports = new SpoListWebhookAddCommand();