import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogScope?: string;
  appCatalogUrl?: string;
  force?: boolean;
  id: string;
}

class SpoAppRemoveCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified app from the specified app catalog';
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
        appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
        force: (!(!args.options.force)).toString(),
        appCatalogScope: args.options.appCatalogScope || 'tenant'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.appCatalogScope) {
          const testScope: string = args.options.appCatalogScope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `appCatalogScope must be either 'tenant' or 'sitecollection' if specified`;
          }

          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
          }
        }

        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.appCatalogUrl) {
          return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

    const removeApp: () => Promise<void> = async (): Promise<void> => {
      try {
        const spoUrl = await spo.getSpoUrl(logger, this.debug);
        const appCatalogUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

        if (this.debug) {
          await logger.logToStderr(`Retrieved app catalog URL ${appCatalogUrl}. Removing app from the app catalog...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(args.options.id)}')/remove`,
          headers: {
            accept: 'application/json;odata=nometadata'
          }
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataPromise(err);
      }
    };

    if (args.options.force) {
      await removeApp();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to remove the app ${args.options.id} from the app catalog?`);

      if (result) {
        await removeApp();
      }
    }
  }
}

export default new SpoAppRemoveCommand();