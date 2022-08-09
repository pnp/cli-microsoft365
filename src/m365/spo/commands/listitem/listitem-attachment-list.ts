import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  itemId: string;
  listId?: string;
  listTitle?: string;
  webUrl: string;
}

class SpoListItemAttachmentListCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ATTACHMENT_LIST;
  }

  public get description(): string {
    return 'Gets the attachments associated to a list item';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listTitle: typeof args.options.listTitle !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--itemId <itemId>'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--listTitle [listTitle]'
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

        if (!args.options.listId && !args.options.listTitle) {
          return `Specify listId or listTitle`;
        }

        if (args.options.listId && args.options.listTitle) {
          return `Specify listId or listTitle but not both`;
        }

        if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
          return `${args.options.listId} in option listId is not a valid GUID`;
        }

        if (isNaN(parseInt(args.options.itemId))) {
          return `${args.options.itemId} is not a number`;
        }

        return true;
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['FileName', 'ServerRelativeUrl'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitleArgument)}')`);

    const requestOptions: any = {
      url: `${listRestUrl}/items(${args.options.itemId})?$select=AttachmentFiles&$expand=AttachmentFiles`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((attachmentFiles: any): void => {
        if (attachmentFiles.AttachmentFiles && attachmentFiles.AttachmentFiles.length > 0) {
          logger.log(attachmentFiles.AttachmentFiles);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No attachments found');
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoListItemAttachmentListCommand();
