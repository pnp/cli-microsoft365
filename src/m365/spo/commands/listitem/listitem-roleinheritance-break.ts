import { Cli } from '../../../../cli/Cli.js';
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
  listItemId: number;
  listId?: string;
  listTitle?: string;
  listUrl?: string;
  clearExistingPermissions?: boolean;
  force?: boolean;
}

class SpoListItemRoleInheritanceBreakCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEINHERITANCE_BREAK;
  }

  public get description(): string {
    return 'Break inheritance of list item';
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
        clearExistingPermissions: args.options.clearExistingPermissions === true,
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
        option: '--listItemId <listItemId>'
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
        option: '-c, --clearExistingPermissions'
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
          return `${args.options.listId} is not a valid GUID`;
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
    if (this.verbose) {
      await logger.logToStderr(`Breaking role inheritance of list item in site at ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.breakListItemRoleInheritance(args.options);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to break the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?`);

      if (result) {
        await this.breakListItemRoleInheritance(args.options);
      }
    }
  }

  private async breakListItemRoleInheritance(options: Options): Promise<void> {
    try {
      let requestUrl: string = `${options.webUrl}/_api/web`;

      if (options.listId) {
        requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')`;
      }
      else if (options.listTitle) {
        requestUrl += `/lists/getbytitle('${formatting.encodeQueryParameter(options.listTitle)}')`;
      }
      else if (options.listUrl) {
        const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
        requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')`;
      }

      let keepExistingPermissions: boolean = true;
      if (options.clearExistingPermissions) {
        keepExistingPermissions = !options.clearExistingPermissions;
      }

      const requestOptions: CliRequestOptions = {
        url: `${requestUrl}/items(${options.listItemId})/breakroleinheritance(${keepExistingPermissions})`,
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListItemRoleInheritanceBreakCommand();