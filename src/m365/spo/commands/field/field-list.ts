import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
}

class SpoFieldListCommand extends SpoCommand {
  public get name(): string {
    return commands.FIELD_LIST;
  }

  public get description(): string {
    return 'Retrieves list of columns for specified list or site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'Group', 'Hidden'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${listUrl}fields`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
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
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-i --listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
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

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    const listOptions: any[] = [args.options.listId, args.options.listTitle, args.options.listUrl];
    if (listOptions.some(item => item !== undefined) && listOptions.filter(item => item !== undefined).length > 1) {
      return `Specify either list id or title or list url`;
    }
    
    return true;
  }
}

module.exports = new SpoFieldListCommand();