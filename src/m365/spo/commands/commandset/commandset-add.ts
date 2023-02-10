import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import request, { CliRequestOptions } from '../../../../request';
import { CustomAction } from '../customaction/customaction';

interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  listType: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
  scope?: string;
  location?: string;
}

class SpoCommandSetAddCommand extends SpoCommand {
  private static readonly listTypes: string[] = ['List', 'Library', 'SitePages'];
  private static readonly scopes: string[] = ['Site', 'Web'];
  private static readonly locations: string[] = ['ContextMenu', 'CommandBar', 'Both'];

  public get name(): string {
    return commands.COMMANDSET_ADD;
  }

  public get description(): string {
    return 'Add a ListView Command Set to a site.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        scope: typeof args.options.scope !== 'undefined',
        location: typeof args.options.location !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-t, --title <title>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-l, --listType <listType>', autocomplete: SpoCommandSetAddCommand.listTypes
      },
      {
        option: '-i, --clientSideComponentId  <clientSideComponentId>'
      },
      {
        option: '--clientSideComponentProperties  [clientSideComponentProperties]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: SpoCommandSetAddCommand.scopes
      },
      {
        option: '--location [location]', autocomplete: SpoCommandSetAddCommand.locations
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId as string)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (SpoCommandSetAddCommand.listTypes.indexOf(args.options.listType) < 0) {
          return `${args.options.listType} is not a valid list type. Allowed values are ${SpoCommandSetAddCommand.listTypes.join(', ')}`;
        }

        if (args.options.scope && SpoCommandSetAddCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpoCommandSetAddCommand.scopes.join(', ')}`;
        }

        if (args.options.location && SpoCommandSetAddCommand.locations.indexOf(args.options.location) < 0) {
          return `${args.options.location} is not a valid location. Allowed values are ${SpoCommandSetAddCommand.locations.join(', ')}`;
        }

        return validation.isValidSharePointUrl(args.options.webUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding ListView Command Set ${args.options.clientSideComponentId} to site '${args.options.webUrl}'...`);
    }

    if (!args.options.scope) {
      args.options.scope = 'Web';
    }

    const location: string | undefined = args.options.location && this.getLocation(args.options.location);
    const listType: string = this.getListTemplate(args.options.listType);

    try {
      const requestBody: any = {
        Title: args.options.title,
        Location: location,
        ClientSideComponentId: args.options.clientSideComponentId,
        RegistrationId: listType
      };

      if (args.options.clientSideComponentProperties) {
        requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
      }

      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/${args.options.scope}/UserCustomActions`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        data: requestBody,
        responseType: 'json'
      };
      const response = await request.post<CustomAction>(requestOptions);

      logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
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

module.exports = new SpoCommandSetAddCommand();