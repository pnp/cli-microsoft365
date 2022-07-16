import { Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
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
}

class SpoListRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores role inheritance on list or library';
  }

  public optionSets(): string[][] | undefined {
    return [
      ['listId', 'listTitle']
    ];
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Restore role inheritance of list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/lists`;

    if (args.options.listId) {
      requestUrl += `(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else {
      requestUrl += `/getbytitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
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
      .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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
      return `${args.options.listId} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new SpoListRoleInheritanceResetCommand();