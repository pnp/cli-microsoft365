import { Cli } from '../../../../cli/Cli';
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
  confirm: boolean;
}

class SpoListRetentionLabelRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_RETENTIONLABEL_REMOVE;
  }

  public get description(): string {
    return 'Clears the retention label on the specified list or library.';
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
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '-l, --listId [listId]'
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
    if (args.options.confirm) {
      await this.removeListRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the retention label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?`
      });

      if (result.continue) {
        await this.removeListRetentionLabel(logger, args);
      }
    }
  }

  private async removeListRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Clears the retention label from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }

    try {
      const listServerRelativeUrl: string = await this.getListServerRelativeUrl(args, logger);
      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: {
          listUrl: listAbsoluteUrl,
          complianceTagValue: '',
          blockDelete: false,
          blockEdit: false,
          syncToItems: false
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListServerRelativeUrl(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.verbose) {
      logger.logToStderr('Getting the list server relative URL');
    }

    if (args.options.listUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
    }

    let listRestUrl = '';
    if (args.options.listId) {
      listRestUrl = `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else {
      listRestUrl = `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/${listRestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listInstance: ListInstance = await request.get<ListInstance>(requestOptions);
    return listInstance.RootFolder.ServerRelativeUrl;
  }
}

module.exports = new SpoListRetentionLabelRemoveCommand();