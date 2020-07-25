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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.id = (!(!args.options.id)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
      cmd.log(`Retrieving information for webhook ${args.options.id} belonging to list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')/Subscriptions('${encodeURIComponent(args.options.id)}')`;
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
      .get(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        cb();
      }, (err: any): void => {
        if (this.verbose) {
          cmd.log('Specified webhook not found');
        }
        this.handleRejectedODataJsonPromise(err, cmd, cb)
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to retrieve webhooks for is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list from which to retrieve the webhook. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list from which to retrieve the webhook. Specify either listId or listTitle but not both'
      },
      {
        option: '-i, --id [id]',
        description: 'ID of the webhook to retrieve'
      },
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

      return true;
    };
  }
}

module.exports = new SpoListWebhookGetCommand();