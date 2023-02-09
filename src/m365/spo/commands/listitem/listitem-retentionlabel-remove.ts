import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

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
    if (this.verbose) {
      logger.logToStderr(`Removing retention label from list item ${args.options.listItemId} from list '${args.options.listId || args.options.listTitle || args.options.listUrl}' in site at ${args.options.webUrl}...`);
    }
    try {
      let url = `${args.options.webUrl}/_api/web`;
      if (args.options.listId) {
        url += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/items(${args.options.listItemId})/SetComplianceTag()`;
      }
      else if (args.options.listTitle) {
        url += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/items(${args.options.listItemId})/SetComplianceTag()`;
      }
      else {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl!);
        url += `/GetList(@a1)/items(@a2)/SetComplianceTag()?@a1='${formatting.encodeQueryParameter(listServerRelativeUrl)}'&@a2='${args.options.listItemId}'`;
      }

      const requestBody = {
        "complianceTag": "",
        "isTagPolicyHold": false,
        "isTagPolicyRecord": false,
        "isEventBasedTag": false,
        "isTagSuperLock": false,
        "isUnlockedAsDefault": false
      };

      const requestOptions: CliRequestOptions = {
        url: url,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListItemRetentionLabelRemoveCommand();