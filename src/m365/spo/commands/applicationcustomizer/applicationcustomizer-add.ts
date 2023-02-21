import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { CustomAction } from '../customaction/customaction';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  title: string;
  webUrl: string;
  clientSideComponentId: string;
  clientSideComponentProperties?: string;
  scope?: string;
}

class SpoApplicationCustomizerAddCommand extends SpoCommand {
  private static readonly scopes: string[] = ['Site', 'Web'];

  public get name(): string {
    return commands.APPLICATIONCUSTOMIZER_ADD;
  }

  public get description(): string {
    return 'Add an application customizer to a site.';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
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
        option: '-i, --clientSideComponentId <clientSideComponentId>'
      },
      {
        option: '--clientSideComponentProperties [clientSideComponentProperties]'
      },
      {
        option: '-s, --scope [scope]', autocomplete: SpoApplicationCustomizerAddCommand.scopes
      }
    );
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        clientSideComponentProperties: typeof args.options.clientSideComponentProperties !== 'undefined',
        scope: typeof args.options.scope !== 'undefined'
      });
    });
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.webUrl) {
          const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
          if (isValidSharePointUrl !== true) {
            return isValidSharePointUrl;
          }
        }

        if (!validation.isValidGuid(args.options.clientSideComponentId)) {
          return `${args.options.clientSideComponentId} is not a valid GUID`;
        }

        if (args.options.clientSideComponentProperties) {
          try {
            JSON.parse(args.options.clientSideComponentProperties);
          }
          catch (e) {
            return `An error has occurred while parsing clientSideComponentProperties: ${e}`;
          }
        }

        if (args.options.scope && SpoApplicationCustomizerAddCommand.scopes.indexOf(args.options.scope) < 0) {
          return `${args.options.scope} is not a valid value for allowedMembers. Valid values are ${SpoApplicationCustomizerAddCommand.scopes.join(', ')}`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding application customizer with title '${args.options.title}' and clientSideComponentId '${args.options.clientSideComponentId}' to the site`);
    }

    const requestBody: any = {
      Title: args.options.title,
      Name: args.options.title,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: args.options.clientSideComponentId
    };

    if (args.options.clientSideComponentProperties) {
      requestBody.ClientSideComponentProperties = args.options.clientSideComponentProperties;
    }

    const scope = args.options.scope || 'Site';

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/${scope}/UserCustomActions`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    await request.post<CustomAction>(requestOptions);
  }
}

module.exports = new SpoApplicationCustomizerAddCommand();