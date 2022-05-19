import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { urlUtil, validation } from '../../../../utils';
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
  scope?: string;
}

class SpoEventreceiverListCommand extends SpoCommand {
  public get name(): string {
    return commands.EVENTRECEIVER_LIST;
  }

  public get description(): string {
    return 'Lists all the event receivers attached to the web, site or list (if any of the list options are filled in)';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.scope = typeof args.options.scope !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let requestUrl = `${args.options.webUrl}/_api/`;
    let listUrl: string = '';

    if (args.options.listId) {
      listUrl = `lists(guid'${encodeURIComponent(args.options.listId)}')/`;
    }
    else if (args.options.listTitle) {
      listUrl = `lists/getByTitle('${encodeURIComponent(args.options.listTitle)}')/`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      listUrl = `GetList('${encodeURIComponent(listServerRelativeUrl)}')/`;
    }

    if (args.options.scope === null || args.options.scope === 'web'){
      requestUrl += `web/${listUrl}eventreceivers`;
    } else {
      requestUrl += 'site/eventreceivers';
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: any[] }>(requestOptions)
      .then((res: any): void => {
        logger.log(res.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listId  [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-s, --scope [scope]',
        autocomplete: ['web', 'site']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
    if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
      return `Specify either list id or title or list url`;
    }

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    if (args.options.scope && ['web', 'site'].indexOf(args.options.scope) === -1) {
      return `${args.options.scope} is not a valid type value. Allowed values web|site.`;
    }

    if (args.options.scope && args.options.scope === 'site' && (args.options.listId || args.options.listUrl || args.options.listTitle)) {
      return 'Scope cannot be set to site when retrieving list event receivers.';
    }

    return true;
  }
}

module.exports = new SpoEventreceiverListCommand();