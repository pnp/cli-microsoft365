import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import { urlUtil, validation } from '../../../../utils';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    telemetryProps.listUrl = typeof args.options.listUrl !== 'undefined';
    telemetryProps.title = typeof args.options.title !== 'undefined';
    telemetryProps.personal = !!args.options.personal;
    telemetryProps.default = !!args.options.default;
    telemetryProps.orderedView = !!args.options.orderedView;
    telemetryProps.paged = !!args.options.paged;
    telemetryProps.rowLimit = typeof args.options.rowLimit !== 'undefined';
    return telemetryProps;
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
      result += `lists(guid'${encodeURIComponent(options.listId)}')`;
    }
    else if (options.listTitle) {
      result += `lists/getByTitle('${encodeURIComponent(options.listTitle)}')`;
    }
    else if (options.listUrl) {
      result += `GetList('${encodeURIComponent(urlUtil.getServerRelativePath(options.webUrl, options.listUrl))}')`;
    }
    result += '/views/add';
    
    return result;
  }

  public optionSets(): string[][] | undefined {
    return [
      ['listId', 'listTitle', 'listUrl']
    ];
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      { option: '-u, --webUrl <webUrl>' },
      { option: '--listId [listId]' },
      { option: '--listTitle [listTitle]' },
      { option: '--listUrl [listUrl]' },
      { option: '--title <title>' },
      { option: '--fields <fields>' },
      { option: '--personal' },
      { option: '--default' },
      { option: '--paged' },
      { option: '--rowLimit [rowLimit]' }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
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
}

module.exports = new SpoListViewAddCommand();