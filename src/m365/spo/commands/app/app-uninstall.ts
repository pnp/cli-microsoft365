import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  siteUrl: string;
  confirm?: boolean;
  appCatalogScope?: string;
}

class SpoAppUninstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UNINSTALL;
  }

  public get description(): string {
    return 'Uninstalls an app from the site';
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
        confirm: (!(!args.options.confirm)).toString(),
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
        option: '-s, --siteUrl <siteUrl>'
      },
      {
        option: '--appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appCatalogScope) {
          const testScope: string = args.options.appCatalogScope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `appCatalogScope must be either 'tenant' or 'sitecollection' if specified`;
          }
        }
    
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }
    
        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const uninstallApp: () => Promise<void> = async (): Promise<void> => {
      const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

      if (this.verbose) {
        logger.logToStderr(`Uninstalling app '${args.options.id}' from the site '${args.options.siteUrl}'...`);
      }

      const requestOptions: any = {
        url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/uninstall`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataPromise(err);
      }
    };

    if (args.options.confirm) {
      await uninstallApp();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to uninstall the app ${args.options.id} from site ${args.options.siteUrl}?`
      });

      if (result.continue) {
        await uninstallApp();
      }
    }
  }
}

module.exports = new SpoAppUninstallCommand();