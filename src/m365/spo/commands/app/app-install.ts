import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  siteUrl: string;
  appCatalogScope?: string;
}

class SpoAppInstallCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_INSTALL;
  }

  public get description(): string {
    return 'Installs an app from the specified app catalog in the site';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

    if (this.verbose) {
      logger.logToStderr(`Installing app '${args.options.id}' in site '${args.options.siteUrl}'...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/install`,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoAppInstallCommand();