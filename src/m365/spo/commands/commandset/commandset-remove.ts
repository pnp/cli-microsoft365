import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import request from '../../../../request';
import { CustomAction } from '../customaction/customaction';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { Cli } from '../../../../cli/Cli';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  scope?: string;
  confirm?: boolean;
}

class SpoCommandsetRemoveCommand extends SpoCommand {
  private static readonly scopes: string[] = ['All', 'Site', 'Web'];

  public get name(): string {
    return commands.COMMANDSET_REMOVE;
  }

  public get description(): string {
    return 'Remove a ListView Command Set that is added to a site.';
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
        title: typeof args.options.title !== 'undefined',
        id: typeof args.options.id !== 'undefined',
        clientSideComponentId: typeof args.options.clientSideComponentId !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        confirm: !!args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-i, --id [id]'
      },
      {
        option: '-c, --clientSideComponentId  [clientSideComponentId]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: SpoCommandsetRemoveCommand.scopes
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id as string)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId as string)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.scope && SpoCommandsetRemoveCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpoCommandsetRemoveCommand.scopes.join(', ')}`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'title', 'clientSideComponentId'] }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Removing ListView Command Set ${args.options.clientSideComponentId || args.options.title || args.options.id} to site '${args.options.webUrl}'...`);
    }

    if (args.options.confirm) {
      await this.deleteCommandset(args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove command set '${args.options.clientSideComponentId || args.options.title || args.options.id}'?`
      });

      if (result.continue) {
        await this.deleteCommandset(args);
      }
    }

  }

  private async getCommandSetId(options: Options): Promise<string> {
    if (options.id) {
      return options.id;
    }
    else if (options.title) {
      const commandSets: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `(Title eq '${formatting.encodeQueryParameter(options.title as string)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);

      if (commandSets.length === 0) {
        throw `No user commandsets with title '${options.title}' found`;
      }

      if (commandSets.length > 1) {
        throw `Multiple user commandsets with title '${options.title}' found. Please disambiguate using IDs: ${commandSets.map(a => a.Id).join(', ')}`;
      }

      return commandSets[0].Id;
    }
    else {
      const commandSets: CustomAction[] = await spo.getCustomActions(options.webUrl, options.scope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);

      if (commandSets.length === 0) {
        throw `No user commandsets with ClientSideComponentId '${options.clientSideComponentId}' found`;
      }
      if (commandSets.length > 1) {
        throw `Multiple user commandsets with ClientSideComponentId '${options.clientSideComponentId}' found. Please disambiguate using IDs: ${commandSets.map((commandSet: CustomAction) => commandSet.Id).join(', ')}`;
      }
      return commandSets[0].Id;
    }
  }

  private async deleteCommandset(args: CommandArgs): Promise<void> {
    if (!args.options.scope) {
      args.options.scope = 'Site';
    }

    try {
      const id = await this.getCommandSetId(args.options);

      const requestOptions: any = {
        url: `${args.options.webUrl}/_api/${args.options.scope}/UserCustomActions('${formatting.encodeQueryParameter(id)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete<CustomAction>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoCommandsetRemoveCommand();