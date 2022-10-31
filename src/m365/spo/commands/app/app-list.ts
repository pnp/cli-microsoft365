import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogScope?: string;
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
        appCatalogScope: args.options.appCatalogScope || 'tenant'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-s, --appCatalogScope [appCatalogScope]',
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
        if (args.options.appCatalogScope) {
          const testScope: string = args.options.appCatalogScope.toLowerCase();

          if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
            return `appCatalogScope must be either 'tenant' or 'sitecollection'`;
          }

          if (testScope === 'sitecollection' && !args.options.appCatalogUrl) {
            return `You must specify appCatalogUrl when the appCatalogScope is sitecollection`;
          }

          if (args.options.appCatalogUrl) {
            return validation.isValidSharePointUrl(args.options.appCatalogUrl);
          }
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';
    let spoUrl: string = '';

    try {
      spoUrl = await spo.getSpoUrl(logger, this.debug);
      appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

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

      const apps = await request.get<{ value: any[] }>(requestOptions);
      if (apps.value && apps.value.length > 0) {
        logger.log(apps.value);
      }
      else {
        if (this.verbose) {
          logger.logToStderr('No apps found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoAppListCommand();