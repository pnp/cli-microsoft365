import { Logger } from '../../../../cli';
import {
  CommandOption,
  CommandTypes
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListItemInstance } from './ListItemInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;
  properties?: string;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listIdArgument)}')` :
      `${args.options.webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitleArgument)}')`);

    const propertiesSelect: string[] = args.options.properties ? args.options.properties.split(',') : [];
    const propertiesWithSlash: string[] = propertiesSelect.filter(item => item.includes('/'));
    const propertiesToExpand: string[] = propertiesWithSlash.map(e => e.split('/')[0]);
    const expandPropertiesArray: string[] = propertiesToExpand.filter((item, pos) => propertiesToExpand.indexOf(item) === pos);
    const fieldExpand: string = expandPropertiesArray.length > 0 ? `&$expand=${expandPropertiesArray.join(",")}` : ``;

    const requestOptions: any = {
      url: `${listRestUrl}/items(${args.options.id})?$select=${encodeURIComponent(propertiesSelect.join(","))}${fieldExpand}`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((response: any): void => {
        delete response['ID'];
        logger.log(<ListItemInstance>response);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-p, --properties [properties]'
      }
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
        'properties'
      ]
    };
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
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
      !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} in option listId is not a valid GUID`;
    }

    if (isNaN(parseInt(args.options.id))) {
      return `${args.options.id} is not a number`;
    }

    return true;
  }
}

module.exports = new SpoListItemGetCommand();
