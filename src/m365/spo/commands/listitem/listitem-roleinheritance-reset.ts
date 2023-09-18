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
  force?: boolean;
}

class SpoListItemRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores the role inheritance of list item, file, or folder';
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
        option: '--listItemId <listItemId>'
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
      await logger.logToStderr(`Restore role inheritance of list item in site at ${args.options.webUrl}...`);
    }

    if (args.options.force) {
      await this.resetListItemRoleInheritance(args.options);
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to reset the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?`);

      if (result) {
        await this.resetListItemRoleInheritance(args.options);
      }
    }
  }

  private async resetListItemRoleInheritance(options: Options): Promise<void> {
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

      const requestOptions: CliRequestOptions = {
        url: `${requestUrl}/items(${options.listItemId})/resetroleinheritance`,
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

export default new SpoListItemRoleInheritanceResetCommand();