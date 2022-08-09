import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  type?: string;
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

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        type: args.options.type,
        secondary: args.options.secondary
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-t, --type [type]',
        autocomplete: SpoSiteRecycleBinItemListCommand.recycleBinItemType.map(item => item.value)
      },
      {
        option: '-s, --secondary'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (typeof args.options.type !== 'undefined' &&
          !SpoSiteRecycleBinItemListCommand.recycleBinItemType.some(item => item.value === args.options.type)) {
          return `${args.options.type} is not a valid value. Allowed values are ${SpoSiteRecycleBinItemListCommand.recycleBinItemType.map(item => item.value).join(', ')}`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all items from recycle bin at ${args.options.siteUrl}...`);
    }

    const state: string = args.options.secondary ? '2' : '1';

    let requestUrl: string = `${args.options.siteUrl}/_api/site/RecycleBin?$filter=(ItemState eq ${state})`;

    if (typeof args.options.type !== 'undefined') {
      const type = SpoSiteRecycleBinItemListCommand.recycleBinItemType.find(item => item.value === args.options.type);
      if (typeof type !== 'undefined') {
        requestUrl += ` and (ItemType eq ${type.id})`;
      }
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
      .then((response): void => {
        logger.log(response.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private static recycleBinItemType: { id: number, value: string }[] = [
    { id: 1, value: 'files' },
    { id: 3, value: 'listItems' },
    { id: 5, value: 'folders' }
  ];
}

module.exports = new SpoSiteRecycleBinItemListCommand();