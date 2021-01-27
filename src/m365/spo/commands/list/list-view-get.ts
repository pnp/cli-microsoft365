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
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  viewId?: string;
  viewTitle?: string;
}

class SpoListViewGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_GET;
  }

  public get description(): string {
    return 'Gets information about specific list view';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.viewId = typeof args.options.viewId !== 'undefined';
    telemetryProps.viewTitle = typeof args.options.viewTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const baseRestUrl: string = `${args.options.webUrl}/_api/web`;
    let listRestUrl: string = '';

    if (args.options.listId) {
      listRestUrl = `/lists(guid'${encodeURIComponent(args.options.listId)}')`;
    }
    else if (args.options.listTitle) {
      listRestUrl = `/lists/getByTitle('${encodeURIComponent(args.options.listTitle as string)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = Utils.getServerRelativePath(args.options.webUrl, args.options.listUrl);

      listRestUrl = `/GetList('${encodeURIComponent(listServerRelativeUrl)}')`;
    }

    const viewRestUrl: string = `/views/${(args.options.viewId ? `getById('${encodeURIComponent(args.options.viewId)}')` : `getByTitle('${encodeURIComponent(args.options.viewTitle as string)}')`)}`;

    const requestOptions: any = {
      url: `${baseRestUrl}${listRestUrl}${viewRestUrl}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((result: any): void => {
        logger.log(result);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--viewId [viewId]'
      },
      {
        option: '--viewTitle [viewTitle]'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.listId && !args.options.listTitle && !args.options.listUrl) {
      return `Specify listId, listTitle or listUrl`;
    }

    if (args.options.listId &&
      !Utils.isValidGuid(args.options.listId)) {
      return `${args.options.listId} in option listId is not a valid GUID`;
    }

    if (!args.options.viewId && !args.options.viewTitle) {
      return `Specify viewId or viewTitle`;
    }

    if (args.options.viewId && args.options.viewTitle) {
      return `Specify viewId or viewTitle but not both`;
    }

    if (args.options.viewId &&
      !Utils.isValidGuid(args.options.viewId)) {
      return `${args.options.viewId} in option viewId is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoListViewGetCommand();