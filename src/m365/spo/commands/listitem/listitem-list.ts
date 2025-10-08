import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  fields?: string;
  filter?: string;
  pageNumber?: number;
  pageSize?: number;
  camlQuery?: string;
  webUrl: string;
}

class SpoListItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_LIST;
  }

  public get description(): string {
    return 'Gets a list of items from the specified list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        fields: typeof args.options.fields !== 'undefined',
        filter: typeof args.options.filter !== 'undefined',
        pageNumber: typeof args.options.pageNumber !== 'undefined',
        pageSize: typeof args.options.pageSize !== 'undefined',
        camlQuery: typeof args.options.camlQuery !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '-s, --pageSize [pageSize]'
      },
      {
        option: '-n, --pageNumber [pageNumber]'
      },
      {
        option: '-q, --camlQuery [camlQuery]'
      },
      {
        option: '--fields [fields]'
      },
      {
        option: '-l, --filter [filter]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.camlQuery && args.options.fields) {
          return `You cannot use the fields-option when specifying a camlQuery`;
        }

        if (args.options.camlQuery && args.options.pageSize) {
          return `You cannot use the pageSize-option when specifying a camlQuery`;
        }

        if (args.options.camlQuery && args.options.pageNumber) {
          return `You cannot use the pageNumber-option when specifying a camlQuery`;
        }

        if (args.options.pageSize && isNaN(Number(args.options.pageSize))) {
          return `pageSize ${args.options.pageSize} must be numeric`;
        }

        if (args.options.pageNumber && !args.options.pageSize) {
          return `pageSize must be specified if pageNumber is specified`;
        }

        if (args.options.pageNumber && isNaN(Number(args.options.pageNumber))) {
          return `pageNumber must be numeric`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] }
    );
  }

  #initTypes(): void {
    this.types.string.push(
      'webUrl',
      'camlQuery',
      'pageSize',
      'pageNumber',
      'fields',
      'filter'
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const options: ListItemListOptions = {
        webUrl: args.options.webUrl,
        listId: args.options.listId,
        listUrl: args.options.listUrl,
        listTitle: args.options.listTitle,
        fields: args.options.fields ? args.options.fields.split(",")
          : (!args.options.output || cli.shouldTrimOutput(args.options.output)) ? ["Title", "Id"] : [],
        filter: args.options.filter,
        pageNumber: args.options.pageNumber,
        pageSize: args.options.pageSize,
        camlQuery: args.options.camlQuery
      };

      const listItems = await spoListItem.getListItems(options, logger, this.verbose);

      await logger.log(listItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemListCommand();