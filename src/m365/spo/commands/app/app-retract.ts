import { Cli, Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  appCatalogUrl?: string;
  scope?: string;
  confirm?: boolean;
}

class SpoAppRetractCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_RETRACT;
  }

  public get description(): string {
    return 'Retracts the specified app from the specified app catalog';
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
        confirm: (!(!args.options.confirm)).toString(),
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
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '-s, --scope [scope]',
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
        if (args.options.scope) {
          const testScope: string = args.options.scope.toLowerCase();
          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `Scope must be either 'tenant' or 'sitecollection' if specified`;
          }
    
          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the scope is sitecollection`;
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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';

    const retractApp: () => void = (): void => {
      spo
        .getSpoUrl(logger, this.debug)
        .then((spoUrl: string): Promise<string> => {
          return this.getAppCatalogSiteUrl(logger, spoUrl, args);
        })
        .then((appCatalogSiteUrl: string): Promise<string> => {
          if (this.verbose) {
            logger.logToStderr(`Retracting app...`);
          }

          const requestOptions: any = {
            url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(args.options.id)}')/retract`,
            headers: {
              accept: 'application/json;odata=nometadata'
            }
          };

          return request.post(requestOptions);
        })
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      retractApp();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to retract the app ${args.options.id} from the app catalog?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          retractApp();
        }
      });
    }
  }
}

module.exports = new SpoAppRetractCommand();