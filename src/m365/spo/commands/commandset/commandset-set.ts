import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import request, { CliRequestOptions } from '../../../../request';
import { CustomAction } from '../customaction/customaction';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  webUrl: string;
  title?: string;
  id?: string;
  clientSideComponentId?: string;
  newClientSideComponentId?: string;
  newTitle: string;
  listType?: string;
  clientSideComponentProperties?: string;
  scope?: string;
  location?: string;
}

class SpoCommandSetSetCommand extends SpoCommand {
  private static readonly listTypes: string[] = ['List', 'Library', 'SitePages'];
  private static readonly scopes: string[] = ['All', 'Site', 'Web'];
  private static readonly locations: string[] = ['ContextMenu', 'CommandBar', 'Both'];

  public get name(): string {
    return commands.COMMANDSET_SET;
  }

  public get description(): string {
    return 'Updates a ListView Command Set on a site.';
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
        newClientSideComponentId: typeof args.options.newClientSideComponentId !== 'undefined',
        listType: typeof args.options.listType !== 'undefined',
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        location: typeof args.options.location !== 'undefined'
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
        option: '--newClientSideComponentId  [newClientSideComponentId]'
      },
      {
        option: '--newTitle [newTitle]'
      },
      {
        option: '-l, --listType [listType]', autocomplete: SpoCommandSetSetCommand.listTypes
      },
      {
        option: '--clientSideComponentProperties  [clientSideComponentProperties]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: SpoCommandSetSetCommand.scopes
      },
      {
        option: '--location [location]', autocomplete: SpoCommandSetSetCommand.locations
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

        if (args.options.newClientSideComponentId && !validation.isValidGuid(args.options.newClientSideComponentId as string)) {
          return `${args.options.newClientSideComponentId} is not a valid GUID`;
        }

        if (args.options.listType && SpoCommandSetSetCommand.listTypes.indexOf(args.options.listType) < 0) {
          return `${args.options.listType} is not a valid list type. Allowed values are ${SpoCommandSetSetCommand.listTypes.join(', ')}`;
        }

        if (args.options.scope && SpoCommandSetSetCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpoCommandSetSetCommand.scopes.join(', ')}`;
        }

        if (args.options.location && SpoCommandSetSetCommand.locations.indexOf(args.options.location) < 0) {
          return `${args.options.location} is not a valid location. Allowed values are ${SpoCommandSetSetCommand.locations.join(', ')}`;
        }

        if (!args.options.newTitle && !args.options.listType && !args.options.clientSideComponentProperties && !args.options.location && !args.options.newClientSideComponentId) {
          return `Please specify option to be updated`;
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
      logger.logToStderr(`Updating ListView Command Set ${args.options.id || args.options.title || args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
    }

    if (!args.options.scope) {
      args.options.scope = 'All';
    }

    const location: string = this.getLocation(args.options.location ? args.options.location : '');

    try {
      const requestBody: any = {};

      if (args.options.newTitle) {
        requestBody.Title = args.options.newTitle;
      }

      if (args.options.location) {
        requestBody.Location = location;
      }

      if (args.options.listType) {
        requestBody.RegistrationId = this.getListTemplate(args.options.listType);
      }

      if (args.options.clientSideComponentProperties) {
        requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
      }

      if (args.options.newClientSideComponentId) {
        requestBody.ClientSideComponentId = args.options.newClientSideComponentId;
      }

      const customAction = await this.getCustomAction(args.options);

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${customAction.Scope === 3 ? "Web" : "Site"}/UserCustomActions('${formatting.encodeQueryParameter(customAction.Id)}')`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'X-HTTP-Method': 'MERGE'
        },
        data: requestBody,
        responseType: 'json'
      };

      await request.post<CustomAction>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getCustomAction(options: Options): Promise<CustomAction> {
    let commandSets: CustomAction[] = [];

    if (options.id) {
      const commandSet = await spo.getCustomActionById(options.webUrl, options.id, options.scope);
      if (commandSet) {
        commandSets.push(commandSet);
      }
    }
    else if (options.title) {
      commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(Title eq '${formatting.encodeQueryParameter(options.title as string)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
    }
    else {
      commandSets = await spo.getCustomActions(options.webUrl, options.scope, `(ClientSideComponentId eq guid'${options.clientSideComponentId}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`);
    }

    if (commandSets.length === 0) {
      throw `No user commandsets with ${options.title && `title '${options.title}'` || options.clientSideComponentId && `ClientSideComponentId '${options.clientSideComponentId}'` || options.id && `id '${options.id}'`} found`;
    }

    if (commandSets.length > 1) {
      throw `Multiple user commandsets with ${options.title ? `title '${options.title}'` : `ClientSideComponentId '${options.clientSideComponentId}'`} found. Please disambiguate using IDs: ${commandSets.map((commandSet: CustomAction) => commandSet.Id).join(', ')}`;
    }

    return commandSets[0];
  }

  private getLocation(location: string): string {
    switch (location) {
      case 'Both':
        return 'ClientSideExtension.ListViewCommandSet';
      case 'ContextMenu':
        return 'ClientSideExtension.ListViewCommandSet.ContextMenu';
      default:
        return 'ClientSideExtension.ListViewCommandSet.CommandBar';
    }
  }

  private getListTemplate(listTemplate: string): string {
    switch (listTemplate) {
      case 'SitePages':
        return '119';
      case 'Library':
        return '101';
      default:
        return '100';
    }
  }
}

module.exports = new SpoCommandSetSetCommand();