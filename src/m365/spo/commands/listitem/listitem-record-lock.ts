import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstance } from '../list/ListInstance';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  listItemId: string;
}

class SpoListItemRecordLockCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RECORD_LOCK;
  }

  public get description(): string {
    return 'Locks the list item record';
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
        listUrl: typeof args.options.listUrl !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listItemId <listItemId>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id: number = parseInt(args.options.listItemId);
        if (isNaN(id)) {
          return `${args.options.listItemId} is not a valid list item ID`;
        }

        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.listId &&
          !validation.isValidGuid(args.options.listId as string)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Locking the list item record ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }
    try {
      let listRestUrl: string = '';
      let listServerRelativeUrl: string = '';

      if (args.options.listUrl) {
        const listServerRelativeUrlFromPath: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);

        listServerRelativeUrl = listServerRelativeUrlFromPath;
      }
      else {
        if (args.options.listId) {
          listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
        }
        else {
          listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
        }

        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/${listRestUrl}?$expand=RootFolder&$select=RootFolder`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const listInstance: ListInstance = await request.get<ListInstance>(requestOptions);
        listServerRelativeUrl = listInstance.RootFolder.ServerRelativeUrl;
      }

      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);
      const requestUrl: string = `${args.options.webUrl}/_api/SP.CompliancePolicy.SPPolicyStoreProxy.LockRecordItem()`;
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          listUrl: listAbsoluteUrl,
          itemId: args.options.listItemId
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListItemRecordLockCommand();