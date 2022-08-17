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
  scope?: string;
}

class SpoAppUpgradeCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_UPGRADE;
  }

  public get description(): string {
    return 'Upgrades app in the specified site';
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    if (this.verbose) {
      logger.logToStderr(`Upgrading app '${args.options.id}' in site '${args.options.siteUrl}'...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/upgrade`,
      headers: {
        accept: 'application/json;odata=nometadata'
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoAppUpgradeCommand();