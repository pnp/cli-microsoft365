import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, urlUtil, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from './ListInstance';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
}

class SpoListSiteScriptGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_SITESCRIPT_GET;
  }

  public get description(): string {
    return 'Extracts a site script from a SharePoint list';
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
        listId: (!(!args.options.listId)).toString(),
        listTitle: (!(!args.options.listTitle)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
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

        if (args.options.listId && args.options.listTitle) {
          return 'Specify listId or listTitle, but not both';
        }

        if (!args.options.listId && !args.options.listTitle) {
          return 'Specify listId or listTitle, one is required';
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      const list: string = (args.options.listId ? args.options.listId : args.options.listTitle) as string;
      logger.logToStderr(`Extracting Site Script from list ${list} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      if (this.debug) {
        logger.logToStderr(`Retrieving List Url from Id '${args.options.listId}'...`);
      }
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder`;
    }
    else {
      if (this.debug) {
        logger.logToStderr(`Retrieving List Url from Title '${args.options.listTitle}'...`);
      }
      requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')?$expand=RootFolder`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<ListInstance>(requestOptions)
      .then((listInstance: ListInstance): Promise<any> => {
        const listAbsoluteUrl = urlUtil.getAbsoluteUrl(args.options.webUrl, listInstance.RootFolder.ServerRelativeUrl);
        const requestUrl = `${args.options.webUrl}/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`;
        const requestOptions: any = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          },
          responseType: 'json',
          data: {
            listUrl: listAbsoluteUrl
          }
        };

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        const siteScript: string | null = res.value;
        if (!siteScript) {
          cb(new CommandError(`An error has occurred, the site script could not be extracted from list '${args.options.listId || args.options.listTitle}'`));
          return;
        }

        logger.log(siteScript);
        cb();
      }, (err: any): void => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }
}

module.exports = new SpoListSiteScriptGetCommand();