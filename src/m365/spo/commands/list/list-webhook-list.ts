import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { formatting } from '../../../../utils/formatting';
import { odata } from '../../../../utils/odata';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listTitle?: string;
  listId?: string;
  listUrl?: string;
}

class SpoListWebhookListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_LIST;
  }

  public get description(): string {
    return 'Lists all webhooks for the specified list';
  }


  public defaultProperties(): string[] | undefined {
    return ['id', 'clientState', 'expirationDateTime', 'resource'];
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
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
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
    if (this.verbose) {
      logger.logToStderr(`Retrieving webhook information for list ${args.options.listTitle || args.options.listId || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions`;
    }

    try {
      const res = await odata.getAllItems<{ id: string, clientState: string, expirationDateTime: Date, resource: string }>(requestUrl);
      if (res && res.length > 0) {
        res.forEach(w => {
          w.clientState = w.clientState || '';
        });

        logger.log(res);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No webhooks found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListWebhookListCommand();