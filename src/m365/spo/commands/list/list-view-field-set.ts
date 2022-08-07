import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  position: string;
  listId?: string;
  listTitle?: string;
  viewId?: string;
  viewTitle?: string;
  webUrl: string;
}

class SpoListViewFieldSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_FIELD_SET;
  }

  public get description(): string {
    return 'Updates existing column in an existing view (eg. move to a specific position).';
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
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        viewId: typeof args.options.viewId !== 'undefined',
        viewTitle: typeof args.options.viewTitle !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined'
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
        option: '--position <position>'
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

        const position: number = parseInt(args.options.position);
        if (isNaN(position)) {
          return `${args.options.position} is not a number`;
        }

        if (args.options.listId && args.options.listTitle) {
          return 'Specify listId or listTitle, but not both';
        }

        if (!args.options.listId && !args.options.listTitle) {
          return 'Specify listId or listTitle, one is required';
        }

        if (args.options.viewId && args.options.viewTitle) {
          return 'Specify viewId or viewTitle, but not both';
        }

        if (!args.options.viewId && !args.options.viewTitle) {
          return 'Specify viewId or viewTitle, one is required';
        }

        if (args.options.id && args.options.title) {
          return 'Specify id or title, but not both';
        }

        if (!args.options.id && !args.options.title) {
          return 'Specify id or title, one is required';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${formatting.encodeQueryParameter(args.options.listId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;

    if (this.verbose) {
      logger.logToStderr(`Getting field ${args.options.id || args.options.title}...`);
    }

    this
      .getField(args.options, listSelector)
      .then((field: { InternalName: string; }): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Moving the field ${args.options.id || args.options.title} in view ${args.options.viewId || args.options.viewTitle} to position ${args.options.position}...`);
        }

        const moveRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;

        const moveRequestOptions: any = {
          url: moveRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: {
            field: field.InternalName,
            index: args.options.position
          },
          responseType: 'json'
        };

        return request.post(moveRequestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

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

module.exports = new SpoListViewFieldSetCommand();