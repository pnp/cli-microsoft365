import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { spo, validation } from '../../../../utils';
import commands from '../../commands';
import { AppMetadata } from './AppMetadata';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  id?: string;
  name?: string;
  scope?: string;
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
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: (!(!args.options.id)).toString(),
        name: (!(!args.options.name)).toString(),
        appCatalogUrl: (!(!args.options.appCatalogUrl)).toString(),
        scope: args.options.scope || 'tenant'
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
        option: '-s, --scope [scope]',
        autocomplete: ['tenant', 'sitecollection']
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
        }

        if (!args.options.id && !args.options.name) {
          return 'Specify either the id or the name';
        }

        if (args.options.id && args.options.name) {
          return 'Specify either the id or the name but not both';
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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';

    try {
      const spoUrl = await spo.getSpoUrl(logger, this.debug);
      appCatalogSiteUrl = await this.getAppCatalogSiteUrl(logger, spoUrl, args);

      if (args.options.id) {
        throw { UniqueId: args.options.id };
      }

      if (this.verbose) {
        logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
      }

      let requestOptions: any = {
        url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const res = await request.get<{ UniqueId: string }>(requestOptions);

      if (this.verbose) {
        logger.logToStderr(`Retrieving information for app ${res}...`);
      }

      requestOptions = {
        url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(res.UniqueId)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const availableApps = request.get<AppMetadata>(requestOptions);
      logger.log(availableApps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoAppGetCommand();