import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.title = (!(!args.options.title)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (args.options.title && this.verbose) {
      logger.log(chalk.yellow(`Option 'title' is deprecated. Please use 'listTitle' instead`));
    }

    if (args.options.id && this.verbose) {
      logger.log(chalk.yellow(`Option 'id' is deprecated. Please use 'listId' instead`));
    }

    if (this.verbose) {
      const list: string = args.options.id ? encodeURIComponent(args.options.id as string) : (args.options.listId ? encodeURIComponent(args.options.listId as string) : (args.options.title ? encodeURIComponent(args.options.title as string) : encodeURIComponent(args.options.listTitle as string)));
      logger.log(`Retrieving webhook information for list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.id) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.id)}')/Subscriptions`;
    }
    else if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions`;
    }
    else if (args.options.listTitle) {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.title as string)}')/Subscriptions`;
    }

    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<{ value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }>(requestOptions)
      .then((res: { value: [{ id: string, clientState: string, expirationDateTime: Date, resource: string }] }): void => {
        if (res.value && res.value.length > 0) {
          if (args.options.output === 'json') {
            logger.log(res.value);
          }
          else {
            logger.log(res.value.map(e => {
              return {
                id: e.id,
                clientState: e.clientState || '',
                expirationDateTime: e.expirationDateTime,
                resource: e.resource
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            logger.log('No webhooks found');
          }
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: `URL of the site where the list to retrieve webhooks for is located`
      },
      {
        option: '-i, --listId [listId]',
        description: `ID of the list to retrieve all webhooks for. Specify either listId or listTitle but not both`
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: `Title of the list to retrieve all webhooks for. Specify either listId or listTitle but not both`
      },
      {
        option: '--id [id]',
        description: `(deprecated. Use 'listId' instead) ID of the list to retrieve all webhooks for. Specify either id or title but not both`
      },
      {
        option: '--title [title]',
        description: `(deprecated. Use 'listTitle' instead) Title of the list to retrieve all webhooks for. Specify either id or title but not both`
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.id) {
      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }
    }

    if (args.options.listId) {
      if (!Utils.isValidGuid(args.options.listId)) {
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
}

module.exports = new SpoListWebhookListCommand();