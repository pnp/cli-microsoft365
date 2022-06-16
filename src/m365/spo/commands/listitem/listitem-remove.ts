import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting, validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
  id: string;
  recycle?: boolean;
  confirm?: boolean;
}

class SpoListItemRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified list item';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = (!(!args.options.listId)).toString();
    telemetryProps.listTitle = (!(!args.options.listTitle)).toString();
    telemetryProps.recycle = (!(!args.options.recycle)).toString();
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const removeListItem: () => void = (): void => {
      if (this.verbose) {
        logger.logToStderr(`Removing list item in site at ${args.options.webUrl}...`);
      }

      let requestUrl: string = '';

      if (args.options.listId) {
        requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else {
        requestUrl = `${args.options.webUrl}/_api/web/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
      }

      requestUrl += `/items(${args.options.id})`;

      if (args.options.recycle) {
        requestUrl += `/recycle()`;
      }

      const requestOptions: any = {
        url: requestUrl,
        method: 'POST',
        headers: {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      request
        .post(requestOptions)
        .then((): void => {
          // REST post call doesn't return anything
          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeListItem();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to ${args.options.recycle ? "recycle" : "remove"} the list item ${args.options.id} from list ${args.options.listId || args.options.listTitle} located in site ${args.options.webUrl}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeListItem();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-l, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      },
      {
        option: '--recycle'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const id: number = parseInt(args.options.id);
    if (isNaN(id)) {
      return `${args.options.id} is not a valid list item ID`;
    }

    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.listId &&
      !validation.isValidGuid(args.options.listId as string)) {
      return `${args.options.listId} is not a valid GUID`;
    }

    if (args.options.listId && args.options.listTitle) {
      return 'Specify id or title, but not both';
    }

    if (!args.options.listId && !args.options.listTitle) {
      return 'Specify id or title';
    }

    return true;
  }
}

module.exports = new SpoListItemRemoveCommand();