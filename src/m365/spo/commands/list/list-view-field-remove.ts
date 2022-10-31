import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  id?: string;
  title?: string;
  confirm?: boolean;
}

class SpoListViewFieldRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified field from list view';
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
        viewId: typeof args.options.viewId !== 'undefined',
        viewTitle: typeof args.options.viewTitle !== 'undefined',
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
        option: '--viewId [viewId]'
      },
      {
        option: '--viewTitle [viewTitle]'
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

        if (args.options.viewId) {
          if (!validation.isValidGuid(args.options.viewId)) {
            return `${args.options.viewId} is not a valid GUID`;
          }
        }

        if (args.options.id) {
          if (!validation.isValidGuid(args.options.id)) {
            return `${args.options.viewId} is not a valid GUID`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['listId', 'listTitle'],
      ['viewId', 'viewTitle'],
      ['id', 'title']
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const listSelector: string = args.options.listId ? `(guid'${formatting.encodeQueryParameter(args.options.listId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;

    const removeFieldFromView: () => Promise<void> = async (): Promise<void> => {
      try {
        if (this.verbose) {
          logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
        }

        const field = await this.getField(args.options, listSelector);
        if (this.verbose) {
          logger.logToStderr(`Removing field ${args.options.id || args.options.title} from view ${args.options.viewId || args.options.viewTitle}...`);
        }

        const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;
        const postRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/removeviewfield('${field.InternalName}')`;

        const postRequestOptions: any = {
          url: postRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(postRequestOptions);
        // REST post call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.confirm) {
      await removeFieldFromView();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the field ${args.options.id || args.options.title} from the view ${args.options.viewId || args.options.viewTitle} from list ${args.options.listId || args.options.listTitle} in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await removeFieldFromView();
      }
    }
  }

  private getField(options: Options, listSelector: string): Promise<{ InternalName: string; }> {
    const fieldSelector: string = options.id ? `/getbyid('${encodeURIComponent(options.id)}')` : `/getbyinternalnameortitle('${encodeURIComponent(options.title as string)}')`;
    const getRequestUrl: string = `${options.webUrl}/_api/web/lists${listSelector}/fields${fieldSelector}`;

    const requestOptions: any = {
      url: getRequestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }
}

module.exports = new SpoListViewFieldRemoveCommand();