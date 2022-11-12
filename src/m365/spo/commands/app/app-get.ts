import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';
import { AppMetadata } from './AppMetadata';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogScope?: string;
  appCatalogUrl?: string;
  id?: string;
  name?: string;
}

class SpoAppGetCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the specified app catalog';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        name: (!(!args.options.name)).toString(),
        appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
        appCatalogScope: args.options.appCatalogScope || 'tenant'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-n, --name [name]'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]'
      },
      {
        option: '-s, --appCatalogScope [appCatalogScope]',
        autocomplete: ['tenant', 'sitecollection']
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
        }

        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.appCatalogUrl) {
          return validation.isValidSharePointUrl(args.options.appCatalogUrl);
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(['id', 'name']);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const scope: string = (args.options.appCatalogScope) ? args.options.appCatalogScope.toLowerCase() : 'tenant';

      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      const appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      let appId: string;
      if (args.options.id) {
        appId = args.options.id;
      }
      else {
        if (this.verbose) {
          logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        const app = await request.get<{ UniqueId: string }>(requestOptions);
        appId = app.UniqueId;
      }

      if (this.verbose) {
        logger.logToStderr(`Retrieving information for app ${appId}...`);
      }

      const requestOptions: any = {
        url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${formatting.encodeQueryParameter(appId)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<AppMetadata>(requestOptions);
      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoAppGetCommand();