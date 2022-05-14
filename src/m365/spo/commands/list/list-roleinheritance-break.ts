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
  clearExistingPermissions?: boolean;
}

class SpoListRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Breaks role inheritance on list or library';
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
    telemetryProps.clearExistingPermissions = args.options.clearExistingPermissions === true;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Breaking role inheritance of list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/lists`;

    if (args.options.listId) {
      requestUrl += `(guid'${encodeURIComponent(args.options.listId)}')`;
    }
    else {
      requestUrl += `/getbytitle('${encodeURIComponent(args.options.listTitle as string)}')`;
    }

    let keepExistingPermissions: boolean = true;
    if (args.options.clearExistingPermissions) {
      keepExistingPermissions = !args.options.clearExistingPermissions;
    }

    const requestOptions: any = {
      url: `${requestUrl}/breakroleinheritance(${keepExistingPermissions})`,
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
      },
      {
        option: '-c, --clearExistingPermissions'
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

module.exports = new SpoListRoleInheritanceBreakCommand();