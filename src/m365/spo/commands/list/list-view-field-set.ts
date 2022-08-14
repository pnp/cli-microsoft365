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
  fieldId?: string;
  fieldTitle?: string;
  fieldPosition: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined',
        viewId: typeof args.options.viewId !== 'undefined',
        viewTitle: typeof args.options.viewTitle !== 'undefined',
        fieldId: typeof args.options.fieldId !== 'undefined',
        fieldTitle: typeof args.options.fieldTitle !== 'undefined'
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
        option: '--fieldId [fieldId]'
      },
      {
        option: '--fieldTitle [fieldTitle]'
      },
      {
        option: '--fieldPosition <fieldPosition>'
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

        if (args.options.fieldId) {
          if (!validation.isValidGuid(args.options.fieldId)) {
            return `${args.options.fieldId} is not a valid GUID`;
          }
        }

        const position: number = parseInt(args.options.fieldPosition);
        if (isNaN(position)) {
          return `${args.options.fieldPosition} is not a number`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      ['listId', 'listTitle'],
      ['viewId', 'viewTitle'],
      ['fieldId', 'fieldTitle']
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listSelector: string = args.options.listId ? `(guid'${formatting.encodeQueryParameter(args.options.listId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
    const viewSelector: string = args.options.viewId ? `('${formatting.encodeQueryParameter(args.options.viewId)}')` : `/GetByTitle('${formatting.encodeQueryParameter(args.options.viewTitle as string)}')`;

    if (this.verbose) {
      logger.logToStderr(`Getting field ${args.options.fieldId || args.options.fieldTitle}...`);
    }

    this
      .getField(args.options, listSelector)
      .then((field: { InternalName: string; }): Promise<void> => {
        if (this.verbose) {
          logger.logToStderr(`Moving the field ${args.options.fieldId || args.options.fieldTitle} in view ${args.options.viewId || args.options.viewTitle} to position ${args.options.fieldPosition}...`);
        }

        const moveRequestUrl: string = `${args.options.webUrl}/_api/web/lists${listSelector}/views${viewSelector}/viewfields/moveviewfieldto`;

        const moveRequestOptions: any = {
          url: moveRequestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          data: {
            field: field.InternalName,
            index: args.options.fieldPosition
          },
          responseType: 'json'
        };

        return request.post(moveRequestOptions);
      })
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));

  }

  private getField(options: Options, listSelector: string): Promise<{ InternalName: string; }> {
    const fieldSelector: string = options.fieldId ? `/getbyid('${encodeURIComponent(options.fieldId)}')` : `/getbyinternalnameortitle('${encodeURIComponent(options.fieldTitle as string)}')`;
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