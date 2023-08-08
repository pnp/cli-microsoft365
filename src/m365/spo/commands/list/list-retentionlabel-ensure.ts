import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListInstance } from './ListInstance.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  name: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  syncToItems?: boolean;
}

class SpoListRetentionLabelEnsureCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_RETENTIONLABEL_ENSURE;
  }

  public get description(): string {
    return 'Sets a default retention label on the specified list or library.';
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
        listUrl: (!(!args.options.listUrl)).toString(),
        syncToItems: args.options.syncToItems || false
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--name <name>'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--syncToItems'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
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
      const requestUrl: string = `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`;
      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          listUrl: listAbsoluteUrl,
          complianceTagValue: args.options.name,
          syncToItems: args.options.syncToItems || false
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

export default new SpoListRetentionLabelEnsureCommand();