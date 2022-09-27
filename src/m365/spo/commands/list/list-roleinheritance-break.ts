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
        clearExistingPermissions: args.options.clearExistingPermissions === true
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
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
          return `${args.options.listId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['listId', 'listTitle']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Breaking role inheritance of list in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/lists`;

    if (args.options.listId) {
      requestUrl += `(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
    }
    else {
      requestUrl += `/getbytitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
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

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListRoleInheritanceBreakCommand();