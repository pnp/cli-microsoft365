import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { spo } from '../../../../utils/spo.js';
import { ListInstance } from './ListInstance.js';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  force: boolean;
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
        force: !!args.options.force
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
        option: '-i, --listId [listId]'
      },
      {
        option: '-l, --listUrl [listUrl]'
      },
      {
        option: '-f, --force'
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
    if (args.options.force) {
      await this.removeListRetentionLabel(logger, args);
    }
    else {
      const result = await Cli.promptForConfirmation({ message: `Are you sure you want to remove the retention label from list '${args.options.listId || args.options.listTitle || args.options.listUrl}'?` });

      if (result) {
        await this.removeListRetentionLabel(logger, args);
      }
    }
  }

  private async removeListRetentionLabel(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Clears the retention label from list ${args.options.listId || args.options.listTitle || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }

    try {
      const listServerRelativeUrl: string = await this.getListServerRelativeUrl(args, logger);
      const listAbsoluteUrl: string = urlUtil.getAbsoluteUrl(args.options.webUrl, listServerRelativeUrl);

      await spo.removeDefaultRetentionLabelFromList(args.options.webUrl, listAbsoluteUrl, logger, args.options.verbose);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getListServerRelativeUrl(args: CommandArgs, logger: Logger): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr('Getting the list server relative URL');
    }

    if (args.options.listUrl) {
      return urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/`;

    if (args.options.listId) {
      requestUrl += `lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/`;
    }
    else {
      requestUrl += `lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')/`;
    }

    const requestOptions: CliRequestOptions = {
      url: `${requestUrl}?$expand=RootFolder&$select=RootFolder/ServerRelativeUrl`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const listInstance = await request.get<ListInstance>(requestOptions);
    return listInstance.RootFolder.ServerRelativeUrl;
  }
}

export default new SpoListRetentionLabelRemoveCommand();