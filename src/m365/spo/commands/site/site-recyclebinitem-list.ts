import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { RecycleBinItemPropertiesCollection } from './SiteProperties';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  type?: number;
  secondary?: boolean;
}

class SpoSiteRecycleBinItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists items from recycle bin';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'DirName'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all items from recycle bin at ${args.options.siteUrl}...`);
    }

    const state: string = args.options.secondary ? '2' : '1';

    let requestUrl: string = `${args.options.siteUrl}/_api/site/RecycleBin?$filter=(ItemState eq ${state})`;

    if (typeof args.options.type !== 'undefined') {
      requestUrl += ` and (ItemType eq ${args.options.type})`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<RecycleBinItemPropertiesCollection>(requestOptions)
      .then((recycleBinItemProperties: RecycleBinItemPropertiesCollection): void => {
        logger.log(recycleBinItemProperties.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-t, --type [type]'
      },
      {
        option: '-s, --secondary'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.siteUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (typeof args.options.type !== 'undefined' && [1, 3, 5].indexOf(args.options.type) < 0) {
      return `${args.options.type} is not a valid value. Allowed values are 1|3|5`;
    }

    return true;
  }
}

module.exports = new SpoSiteRecycleBinItemListCommand();