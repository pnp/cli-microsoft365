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
  confirm?: boolean;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  webUrl: string;
}

class SpoListViewRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_REMOVE;
  }

  public get description(): string {
    return 'Deletes the specified view from the list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeViewFromList: () => void = (): void => {
      if (this.verbose) {
        const list: string = args.options.listId ? encodeURIComponent(args.options.listId as string) : encodeURIComponent(args.options.listTitle as string);
        cmd.log(`Removing view ${args.options.viewId || args.options.viewTitle} from list ${list} in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';
      const listSelector: string = args.options.listId ? `(guid'${encodeURIComponent(args.options.listId)}')` : `/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;
      const viewSelector: string = args.options.viewId ? `(guid'${encodeURIComponent(args.options.viewId)}')` : `/GetByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`;

      requestUrl = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}`;

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };

      request
        .post(requestOptions)
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeViewFromList();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the view ${args.options.viewId || args.options.viewTitle} from the list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeViewFromList();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list is located'
      },
      {
        option: '--listId [listId]',
        description: 'ID of the list from which to remove the view. Specify listId or listTitle but not both'
      },
      {
        option: '--listTitle [listTitle]',
        description: 'Title of the list from which to remove the view. Specify listId or listTitle but not both'
      },
      {
        option: '--viewId [viewId]',
        description: 'ID of the view to remove. Specify viewId or viewTitle but not both'
      },
      {
        option: '--viewTitle [viewTitle]',
        description: 'ID of the view to remove. Specify viewId or viewTitle but not both'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the view from the list'
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

      if (args.options.viewId) {
        if (!Utils.isValidGuid(args.options.viewId)) {
          return `${args.options.viewId} is not a valid GUID`;
        }
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify listId or listTitle, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify listId or listTitle, one is required';
      }

      if (args.options.viewId && args.options.viewTitle) {
        return 'Specify viewId or viewTitle, but not both';
      }

      if (!args.options.viewId && !args.options.viewTitle) {
        return 'Specify viewId or viewTitle, one is required';
      }

      return true;
    };
  }
}

module.exports = new SpoListViewRemoveCommand();