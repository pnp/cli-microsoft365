import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
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
  title: string;
  fields: string;
  viewQuery?: string;
  personal?: boolean;
  default?: boolean;
  paged?: boolean;
  rowLimit?: number;
}

class SpoListViewAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_VIEW_ADD;
  }

  public get description(): string {
    return 'Adds a new view to a SharePoint list.';
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
        title: typeof args.options.title !== 'undefined',
        viewQuery: typeof args.options.viewQuery !== 'undefined',
        personal: !!args.options.personal,
        default: !!args.options.default,
        orderedView: !!args.options.orderedView,
        paged: !!args.options.paged,
        rowLimit: typeof args.options.rowLimit !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-u, --webUrl <webUrl>' },
      { option: '--listId [listId]' },
      { option: '--listTitle [listTitle]' },
      { option: '--listUrl [listUrl]' },
      { option: '--title <title>' },
      { option: '--fields <fields>' },
      { option: '--viewQuery [viewQuery]' },
      { option: '--personal' },
      { option: '--default' },
      { option: '--paged' },
      { option: '--rowLimit [rowLimit]' }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const webUrlValidation = validation.isValidSharePointUrl(args.options.webUrl);
        if (webUrlValidation !== true) {
          return webUrlValidation;
        }
    
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }
    
        if (args.options.rowLimit !== undefined) {
          if (isNaN(args.options.rowLimit)) {
            return `${args.options.rowLimit} is not a number`;
          }
    
          if (+args.options.rowLimit <= 0) {
            return 'rowLimit option must be greater than 0.';
          }
        }
    
        if (args.options.personal && args.options.default) {
          return 'Default view cannot be a personal view.';
        }
    
        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listTitle', 'listUrl']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: AxiosRequestConfig = {
      url: this.getRestUrl(args.options),
      headers: {
        'content-type': 'application/json;odata=verbose',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        parameters: {
          Title: args.options.title,
          ViewFields: {
            results: args.options.fields.split(',')
          },
          Query: args.options.viewQuery,
          PersonalView: !!args.options.personal,
          SetAsDefaultView: !!args.options.default,
          Paged: !!args.options.paged,
          RowLimit: args.options.rowLimit ? +args.options.rowLimit : 30
        }
      }
    };

    request
      .post(requestOptions)
      .then((result: any): void => {
        logger.log(result);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  private getRestUrl(options: Options): string {
    let result: string = `${options.webUrl}/_api/web/`;
    if (options.listId) {
      result += `lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      result += `lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      result += `GetList('${formatting.encodeQueryParameter(urlUtil.getServerRelativePath(options.webUrl, options.listUrl))}')`;
    }
    result += '/views/add';

    return result;
  }
}

module.exports = new SpoListViewAddCommand();