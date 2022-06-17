import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  scope?: string;
  appCatalogUrl?: string;
}

class SpoAppListCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the specified app catalog';
  }

  public defaultProperties(): string[] | undefined {
    return [`Title`, `ID`, `Deployed`, `AppCatalogVersion`];
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
        scope: args.options.scope || 'tenant'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --scope [scope]',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        // verify either 'tenant' or 'sitecollection' specified if scope provided
        if (args.options.scope) {
          const testScope: string = args.options.scope.toLowerCase();

          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `Scope must be either 'tenant' or 'sitecollection'`;
          }

          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the scope is sitecollection`;
          }

          if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
          }
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';
    let spoUrl: string = '';

    spo
      .getSpoUrl(logger, this.debug)
      .then((_spoUrl: string): Promise<string> => {
        spoUrl = _spoUrl;
        return this.getAppCatalogSiteUrl(logger, spoUrl, args);
      })
      .then((appCatalogUrl: string): Promise<{ value: any[] }> => {
        appCatalogSiteUrl = appCatalogUrl;

        if (this.verbose) {
          logger.logToStderr(`Retrieving apps...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((apps: { value: any[] }): void => {
        if (apps.value && apps.value.length > 0) {
          logger.log(apps.value);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No apps found');
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }
}

module.exports = new SpoAppListCommand();