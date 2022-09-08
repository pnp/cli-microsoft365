import { Cli, Logger } from '../../../../cli';
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
  listItemId: number;
  listId?: string;
  listTitle?: string;
  confirm?: boolean;
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
        confirm: (!(!args.options.confirm)).toString()
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
        option: '--confirm'
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
    this.optionSets.push(['listId', 'listTitle']);
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Restore role inheritance of list item in site at ${args.options.webUrl}...`);
    }

    const resetListItemRoleInheritance = (): void => {
      let requestUrl: string = `${args.options.webUrl}/_api/web/lists`;

      if (args.options.listId) {
        requestUrl += `(guid'${formatting.encodeQueryParameter(args.options.listId)}')`;
      }
      else {
        requestUrl += `/getbytitle('${formatting.encodeQueryParameter(args.options.listTitle as string)}')`;
      }
  
      const requestOptions: any = {
        url: `${requestUrl}/items(${args.options.listItemId})/resetroleinheritance`,
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
    };

    if (args.options.confirm) {
      resetListItemRoleInheritance();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to reset the role inheritance of ${args.options.listItemId} in list ${args.options.listId ?? args.options.listTitle}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          resetListItemRoleInheritance();
        }
      });
    }
  }
}

module.exports = new SpoListItemRoleInheritanceResetCommand();