import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['FileName', 'ServerRelativeUrl'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const listIdArgument = args.options.listId || '';
    const listTitleArgument = args.options.listTitle || '';
    const listRestUrl: string = (args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`);

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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (!args.options.listId && !args.options.listTitle) {
      return `Specify listId or listTitle`;
    }

    if (args.options.listId && args.options.listTitle) {
      return `Specify listId or listTitle but not both`;
    }

    if (args.options.listId && !Utils.isValidGuid(args.options.listId)) {
      return `${args.options.listId} in option listId is not a valid GUID`;
    }

    if (isNaN(parseInt(args.options.itemId))) {
      return `${args.options.itemId} is not a number`;
    }

    return true;
  }
}

module.exports = new SpoListItemAttachmentListCommand();
