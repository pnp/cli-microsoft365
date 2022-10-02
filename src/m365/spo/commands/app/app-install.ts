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
  scope?: string;
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
        scope: args.options.scope || 'tenant'
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
        option: '--scope [scope]',
        autocomplete: ['tenant', 'sitecollection']
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.scope) {
          const testScope: string = args.options.scope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `Scope must be either 'tenant' or 'sitecollection' if specified`;
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
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    if (this.verbose) {
      logger.logToStderr(`Installing app '${args.options.id}' in site '${args.options.siteUrl}'...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/install`,
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
  }
}

module.exports = new SpoAppInstallCommand();