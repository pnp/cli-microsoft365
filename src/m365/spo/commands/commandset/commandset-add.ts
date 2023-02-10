import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import request from '../../../../request';
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

class SpoCommandsetAddCommand extends SpoCommand {
  private static readonly listTypes: string[] = ['List', 'Library'];
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
        option: '-l, --listType <listType>', autocomplete: SpoCommandsetAddCommand.listTypes
      },
      {
        option: '-i, --clientSideComponentId  <clientSideComponentId>'
      },
      {
        option: '--clientSideComponentProperties  [clientSideComponentProperties]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: SpoCommandsetAddCommand.scopes
      },
      {
        option: '--location [location]', autocomplete: SpoCommandsetAddCommand.locations
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.clientSideComponentId && !validation.isValidGuid(args.options.clientSideComponentId as string)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (SpoCommandsetAddCommand.listTypes.indexOf(args.options.listType) < 0) {
          return `${args.options.listType} is not a valid list type. Allowed values are ${SpoCommandsetAddCommand.listTypes.join(', ')}`;
        }

        if (args.options.scope && SpoCommandsetAddCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid scope. Allowed values are ${SpoCommandsetAddCommand.scopes.join(', ')}`;
        }

        if (args.options.location && SpoCommandsetAddCommand.locations.indexOf(args.options.location) < 0) {
          return `${args.options.location} is not a valid location. Allowed values are ${SpoCommandsetAddCommand.locations.join(', ')}`;
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
      args.options.scope = 'Site';
    }

    const location: string = this.getLocation(args.options.location ? args.options.location : '');

    try {
      const requestBody: any = {
        Title: args.options.title,
        Location: location,
        ClientSideComponentId: args.options.clientSideComponentId,
        RegistrationId: args.options.listType === 'List' ? "100" : "101"
      };

      if (args.options.clientSideComponentProperties) {
        requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
      }

      const requestOptions: any = {
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
}

module.exports = new SpoCommandsetAddCommand();