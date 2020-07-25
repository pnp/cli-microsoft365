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
import * as chalk from 'chalk';

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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.notificationUrl = (!(!args.options.notificationUrl)).toString();
    telemetryProps.expirationDateTime = (!(!args.options.expirationDateTime)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Updating webhook ${args.options.id} belonging to list ${args.options.listId ? encodeURIComponent(args.options.listId) : encodeURIComponent(args.options.listTitle as string)} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
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
      body: requestBody,
      json: true
    };

    request
      .patch(requestOptions)
      .then((): void => {
        // REST patch call doesn't return anything
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list which contains the webhook is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list which contains the webhook which should be updated. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list which contains the webhook which should be updated. Specify either listId or listTitle but not both'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the webhook to update'
      },
      {
        option: '-n, --notificationUrl [notificationUrl]',
        description: 'The new notification url'
      },
      {
        option: '-e, --expirationDateTime [expirationDateTime]',
        description: 'The new expiration date'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

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

      if (!args.options.notificationUrl && !args.options.expirationDateTime) {
        return 'Specify notificationUrl, expirationDateTime or both, at least one is required';
      }

      const parsedDateTime = Date.parse(args.options.expirationDateTime as string)
      if (args.options.expirationDateTime && !(!parsedDateTime) !== true) {
        if (args.options.output === 'json') {
          return `${args.options.expirationDateTime} is not a valid date format. Provide the date in one of the following formats: YYYY-MM-DD, YYYY-MM-DDThh:mm, YYYY-MM-DDThh:mmZ, YYYY-MM-DDThh:mm±hh:mm`;
        }
        else {
          return `${args.options.expirationDateTime} is not a valid date format. Provide the date in one of the following formats:
  ${chalk.grey('YYYY-MM-DD')}
  ${chalk.grey('YYYY-MM-DDThh:mm')}
  ${chalk.grey('YYYY-MM-DDThh:mmZ')}
  ${chalk.grey('YYYY-MM-DDThh:mm±hh:mm')}`;
        }
      }

      return true;
    };
  }
}

module.exports = new SpoListWebhookSetCommand();