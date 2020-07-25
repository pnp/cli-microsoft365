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
  confirm?: boolean;
}

class SpoListWebhookRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified webhook from the list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);

    const removeWebhook: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Webhook ${args.options.id} is about to be removed from list ${list} located at site ${args.options.webUrl}...`);
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
        method: 'DELETE',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .delete(requestOptions)
        .then((): void => {
          // REST delete call doesn't return anything
          cb();
        }, (err: any): void => {
          this.handleRejectedODataJsonPromise(err, cmd, cb)
        });
    }

    if (args.options.confirm) {
      removeWebhook();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove webhook ${args.options.id} from list ${list} located at site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeWebhook();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list to remove the webhook from is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list from which the webhook should be removed. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list from which the webhook should be removed. Specify either listId or listTitle but not both'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the webhook to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the webhook'
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

      return true;
    };
  }
}

module.exports = new SpoListWebhookRemoveCommand();