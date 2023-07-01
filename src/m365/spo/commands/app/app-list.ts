import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
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
        await logger.logToStderr(`Retrieving apps...`);
      }

      const apps = await odata.getAllItems<any>(`${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps`);
      if (apps && apps.length > 0) {
        await logger.log(apps);
      }
      else {
        if (this.verbose) {
          await logger.logToStderr('No apps found');
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppListCommand();