import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm?: boolean;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  id?: string;
  title?: string;
  webUrl: string;
}

class SpoListViewRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_REMOVE;
  }

  public get description(): string {
    return 'Deletes the specified view from the list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        listUrl: typeof args.options.listUrl !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        confirm: (!(!args.options.confirm)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
        option: '--id [id]'
      },
      {
        option: '--title [title]'
      },
      {
        option: '--confirm'
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

        if (args.options.listId) {
          if (!validation.isValidGuid(args.options.listId)) {
            return `${args.options.listId} is not a valid GUID`;
          }
        }

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.id} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['listId', 'listTitle', 'listUrl'],
      ['id', 'title']
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeViewFromList: () => Promise<void> = async (): Promise<void> => {
      let listSelector: string = '';
      if (args.options.listId) {
        listSelector = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        listSelector = `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        listSelector = `GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const viewSelector: string = args.options.id ? `(guid'${formatting.encodeQueryParameter(args.options.id)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.title as string)}')`;

      const requestUrl: string = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}`;

      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      try {
        await request.post(requestOptions);
        // REST post call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeViewFromList();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the view ${args.options.id || args.options.title} from the list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await removeViewFromList();
      }
    }
  }
}

module.exports = new SpoListViewRemoveCommand();