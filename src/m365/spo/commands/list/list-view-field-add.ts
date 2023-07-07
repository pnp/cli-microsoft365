import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  position?: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  viewId?: string;
  viewTitle?: string;
  webUrl: string;
}

class SpoListViewFieldAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_ADD;
  }

  public get description(): string {
    return 'Adds the specified field to list view';
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
        viewId: typeof args.options.viewId !== 'undefined',
        viewTitle: typeof args.options.viewTitle !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        position: typeof args.options.position !== 'undefined'
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
        option: '--position [position]'
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
            return `${args.options.id} is not a valid GUID`;
          }
        }

        if (args.options.position) {
          const position: number = parseInt(args.options.position);
          if (isNaN(position)) {
            return `${args.options.position} is not a number`;
          }
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['listId', 'listTitle', 'listUrl'] },
      { options: ['viewId', 'viewTitle'] },
      { options: ['id', 'title'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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

    let viewSelector: string = '';
    let currentField: { InternalName: string; };

    if (this.verbose) {
      logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
    }

    try {
      const field = await this.getField(args.options, listSelector);

      if (this.verbose) {
        logger.logToStderr(`Adding the field ${args.options.id || args.options.title} to the view ${args.options.viewId || args.options.viewTitle}...`);
      }

      currentField = field;

      viewSelector = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;
      const postRequestUrl: string = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}/viewfields/addviewfield('${field.InternalName}')`;

      const postRequestOptions: CliRequestOptions = {
        url: postRequestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(postRequestOptions);

      if (typeof args.options.position === 'undefined') {
        if (this.debug) {
          logger.logToStderr(`No field position.`);
        }

        return;
      }

      if (this.debug) {
        logger.logToStderr(`moveField request...`);
        logger.logToStderr(args.options.position);
      }

      if (this.verbose) {
        logger.logToStderr(`Moving the field ${args.options.id || args.options.title} to the position ${args.options.position} from view ${args.options.viewId || args.options.viewTitle}...`);
      }
      const moveRequestUrl: string = `${args.options.webUrl}/_api/web/${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;

      const moveRequestOptions: CliRequestOptions = {
        url: moveRequestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: { 'field': currentField.InternalName, 'index': args.options.position },
        responseType: 'json'
      };

      await request.post(moveRequestOptions);
      // REST post call doesn't return anything
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getField(options: Options, listSelector: string): Promise<{ InternalName: string; }> {
    const fieldSelector: string = options.id ? `/getbyid('${formatting.encodeQueryParameter(options.id)}')` : `/getbyinternalnameortitle('${formatting.encodeQueryParameter(options.title as string)}')`;
    const getRequestUrl: string = `${options.webUrl}/_api/web/${listSelector}/fields${fieldSelector}`;

    const requestOptions: CliRequestOptions = {
      url: getRequestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }
}

module.exports = new SpoListViewFieldAddCommand();