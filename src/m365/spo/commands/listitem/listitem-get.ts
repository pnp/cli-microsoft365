import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate,
  CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ListItemInstance } from './ListItemInstance';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;
  fields?: string;
}

class SpoListItemGetCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.LISTITEM_GET;
  }

  public get description(): string {
    return 'Gets a list item from the specified list';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

    const fieldSelect: string = args.options.fields ?
      `?$select=${encodeURIComponent(args.options.fields)}` :
      (
        (!args.options.output || args.options.output === 'text') ?
          `?$select=Id,Title` :
          ``
      )

    const requestOptions: any = {
      url: `${listRestUrl}/items(${args.options.id})${fieldSelect}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((response: any): void => {
        (!args.options.output || args.options.output === 'text') && delete response["ID"];
        cmd.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site from which the item should be retrieved'
      },
      {
        option: '-i, --id <id>',
        description: 'ID of the item to retrieve.'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list from which to retrieve the item. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list from which to retrieve the item. Specify listId or listTitle but not both'
      },
      {
        option: '-f, --fields [fields]',
        description: 'Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public types(): CommandTypes {
    return {
      string: [
        'webUrl',
        'listId',
        'listTitle',
        'id',
        'fields',
      ]
    };
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.listId && !args.options.listTitle) {
        return `Specify listId or listTitle`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (isNaN(parseInt(args.options.id))) {
        return `${args.options.id} is not a number`;
      }

      return true;
    };
  }

}

module.exports = new SpoListItemGetCommand();
