import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import * as url from 'url';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  listItemId: string;
  confirm?: boolean;
}

class SpoListItemRetentionLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clear the retention label from a list item';
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
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listItemId <listItemId>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--confirm'
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
    if (args.options.confirm) {
      await this.removeListItemRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the retentionlabel from list item ${args.options.listItemId} from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' located in site ${args.options.webUrl}?`
      });

      if (result.continue) {
        await this.removeListItemRetentionLabel(logger, args);
      }
    }
  }

  private async removeListItemRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const listAbsoluteUrl = await this.getListAbsoluteUrl(args.options, logger);

      await spo.removeRetentionLabel(args.options.webUrl, listAbsoluteUrl, [parseInt(args.options.listItemId)], logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListAbsoluteUrl(options: Options, logger: Logger): Promise<string> {
    const parsedUrl = url.parse(options.webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    if (options.listUrl) {
      const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      return urlUtil.urlCombine(tenantUrl, serverRelativePath);
    }

    if (this.verbose) {
      logger.logToStderr(`Retrieving list absolute URL...`);
    }

    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
    }

    requestUrl += "?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl";

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<{ RootFolder: { ServerRelativeUrl: string } }>(requestOptions);
    const serverRelativePath = urlUtil.getServerRelativePath(options.webUrl, response.RootFolder.ServerRelativeUrl);
    const listAbsoluteUrl = urlUtil.urlCombine(tenantUrl, serverRelativePath);

    if (this.verbose) {
      logger.logToStderr(`List absolute URL found: '${listAbsoluteUrl}'`);
    }

    return listAbsoluteUrl;
  }
}

module.exports = new SpoListItemRetentionLabelRemoveCommand();