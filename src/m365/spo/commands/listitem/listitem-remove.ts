import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoListItemRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list item';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeListItem: () => void = (): void => {
      if (this.verbose) {
        cmd.log(`Removing list item in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.listId) {
        requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')`;
      }
      else {
        requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;
      }

      requestUrl += `/items(${args.options.id})`;

      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }

      const requestOptions: any = {
        url: requestUrl,
        method: 'POST',
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
      removeListItem();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the list item ${args.options.id} from list ${args.options.listId || args.options.listTitle} located in site ${args.options.webUrl}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeListItem();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list item to remove is located'
      },
      {
        option: '-i, --id <id>',
        description: 'The ID of the list item to remove.'
      },
      {
        option: '-l, --listId [listId]',
        description: 'The ID of the list to remove the item from. Specify either listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list to remove the item from. Specify either listId or listTitle but not both'
      },
      {
        option: '--recycle',
        description: 'Recycle the list item instead of actually deleting it'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the list item'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      const id: number = parseInt(args.options.id);
      if (isNaN(id)) {
        return `${args.options.id} is not a valid list item ID`;
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId as string)) {
        return `${args.options.listId} is not a valid GUID`;
      }

      if (args.options.listId && args.options.listTitle) {
        return 'Specify id or title, but not both';
      }

      if (!args.options.listId && !args.options.listTitle) {
        return 'Specify id or title';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Remove the list item with ID ${chalk.grey(1)} from list with ID
    ${chalk.grey('0cd891ef-afce-4e55-b836-fce03286cccf')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
      ${commands.LISTITEM_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1

    Remove the list item with with ID ${chalk.grey(1)} from list with title
    ${chalk.grey('List 1')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} 
      ${commands.LISTITEM_REMOVE} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
      `);
  }
}

module.exports = new SpoListItemRemoveCommand();