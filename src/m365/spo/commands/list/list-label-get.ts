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

class SpoListLabelGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LABEL_GET;
  }

  public get description(): string {
    return 'Gets label set on the specified list';
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
    this.optionSets.push(['listId', 'listTitle', 'listUrl']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr(`Getting label set on the list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.listId) {
        if (this.debug) {
          logger.logToStderr(`Retrieving List from Id '${args.options.listId}'...`);
        }

        requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')?$expand=RootFolder&$select=RootFolder`;
      }
      else if (args.options.listTitle) {
        if (this.debug) {
          logger.logToStderr(`Retrieving List from Title '${args.options.listTitle}'...`);
        }

        requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')?$expand=RootFolder&$select=RootFolder`;
      }
      else if (args.options.listUrl) {
        if (this.debug) {
          logger.logToStderr(`Retrieving List from URL '${args.options.listUrl}'...`);
        }

        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl = `${args.options.webUrl}/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      let requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const listInstance = await request.get<ListInstance>(requestOptions);
      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listInstance.RootFolder.ServerRelativeUrl);
      requestOptions = {
        url: `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_GetListComplianceTag`,
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json',
        data: {
          listUrl: listAbsoluteUrl
        }
      };

      const res = await request.post<any>(requestOptions);
      if (res['odata.null'] !== true) {
        logger.log(res);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListLabelGetCommand();