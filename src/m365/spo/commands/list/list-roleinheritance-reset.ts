import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  listId?: string;
  listTitle?: string;
}

class SpoListRoleInHeritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restore role inheritance on list or library.';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`reset role inheritance of list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = '';

    if (args.options.listId) {
      requestUrl = `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(args.options.listId)}')`;
    }
    else {
      requestUrl = `${args.options.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(args.options.listTitle as string)}')`;
    }

    const requestOptions: any = {
      url: `${requestUrl}/resetroleinheritance`,
      method: 'POST',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json'
      },
      responseType: 'json'
    };

    request
      .post(requestOptions)
      .then((): void => { cb(); }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-i, --listId [listId]'
      },
      {
        option: '-t, --listTitle [listTitle]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.listId && !validation.isValidGuid(args.options.listId)) {
      return `${args.options.id} is not a valid GUID`;
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

module.exports = new SpoListRoleInHeritanceResetCommand();