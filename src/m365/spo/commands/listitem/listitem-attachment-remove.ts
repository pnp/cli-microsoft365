import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  listItemId: number;
  fileName: string;
  force?: boolean;
}

class SpoListItemAttachmentRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ATTACHMENT_REMOVE;
  }

  public get description(): string {
    return 'Removes an attachment from a list item';
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
      },
      {
        option: '--listUrl [listUrl]'
      },
      {
        option: '--listItemId <listItemId>'
      },
      {
        option: '-n, --fileName <fileName>'
      },
      {
        option: '-f, --force'
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

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (isNaN(args.options.listItemId)) {
          return `${args.options.listItemId} is not a number`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listTitle', 'listUrl'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeListItemAttachment = async (): Promise<void> => {
      if (this.verbose) {
        const list: string = (args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl) as string;
        await logger.logToStderr(`Removing attachment ${args.options.fileName} of item with id ${args.options.listItemId} from list ${list} in site at ${args.options.webUrl}...`);
      }

      let requestUrl = `${args.options.webUrl}/_api/web`;

      if (args.options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else if (args.options.listTitle) {
        requestUrl += `/lists/getByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')`;
      }
      else if (args.options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: `${requestUrl}/items(${args.options.listItemId})/AttachmentFiles('${args.options.fileName}')`,
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeListItemAttachment();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the attachment ${args.options.fileName} of item with id ${args.options.listItemId} from the list ${args.options.listId ? args.options.listId : args.options.listTitle ? args.options.listTitle : args.options.listUrl} in site ${args.options.webUrl}?` });

      if (result) {
        await removeListItemAttachment();
      }
    }
  }
}

export default new SpoListItemAttachmentRemoveCommand();