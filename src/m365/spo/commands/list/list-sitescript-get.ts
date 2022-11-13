import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
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
  listUrl?: string;
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
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: (!(!args.options.listId)).toString(),
        listTitle: (!(!args.options.listTitle)).toString(),
        listUrl: (!(!args.options.listUrl)).toString()
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
      },
      {
        option: '--listUrl [listUrl]'
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

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Extracting Site Script from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
      }

      let listServerRelativeUrl: string = '';

      if (args.options.listUrl) {
        if (this.debug) {
          logger.logToStderr(`Retrieving List from URL '${args.options.listUrl}'...`);
        }

        listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      }
      else {
        let requestUrl: string = `${args.options.webUrl}/_api/web/`;

        if (args.options.listId) {
          if (this.debug) {
            logger.logToStderr(`Retrieving List from Id '${args.options.listId}'...`);
          }
          requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder`;
        }
        else if (args.options.listTitle) {
          if (this.debug) {
            logger.logToStderr(`Retrieving List from Title '${args.options.listTitle}'...`);
          }
          requestUrl += `lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')?$expand=RootFolder`;
        }

        const requestOptions: AxiosRequestConfig = {
          url: requestUrl,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const listInstance = await request.get<ListInstance>(requestOptions);
        listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
      }

      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
      const reqOptions: AxiosRequestConfig = {
        url: `${args.options.webUrl}/_api/Microsoft_SharePoint_Utilities_WebTemplateExtensions_SiteScriptUtility_GetSiteScriptFromList`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          listUrl: listAbsoluteUrl
        }
      };

      const res = await request.post<any>(reqOptions);
      const siteScript: string | null = res.value;
      if (!siteScript) {
        throw `An error has occurred, the site script could not be extracted from list '${args.options.listId || args.options.listTitle}'`;
      }

      logger.log(siteScript);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListSiteScriptGetCommand();