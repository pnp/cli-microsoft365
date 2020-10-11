import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.name = (!(!args.options.name)).toString();
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.scope = args.options.scope || 'tenant';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogSiteUrl: string = '';

    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(logger, spoUrl, args);
      })
      .then((appCatalogUrl: string): Promise<{ UniqueId: string }> => {
        appCatalogSiteUrl = appCatalogUrl;

        if (args.options.id) {
          return Promise.resolve({ UniqueId: args.options.id });
        }

        if (this.verbose) {
          logger.log(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: { UniqueId: string }): Promise<AppMetadata> => {
        if (this.verbose) {
          logger.log(`Retrieving information for app ${res}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(res.UniqueId)}')`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })
      .then((res: AppMetadata): void => {
        logger.log(res);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant or site collection app catalog. It must be specified when the scope is \'sitecollection\''
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    // verify either 'tenant' or 'sitecollection' specified if scope provided
    if (args.options.scope) {
      const testScope: string = args.options.scope.toLowerCase();
      if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
        return `Scope must be either 'tenant' or 'sitecollection'`
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

    if (args.options.id && !Utils.isValidGuid(args.options.id)) {
      return `${args.options.id} is not a valid GUID`;
    }

    if (args.options.appCatalogUrl) {
      return SpoAppBaseCommand.isValidSharePointUrl(args.options.appCatalogUrl);
    }

    return true;
  }
}

module.exports = new SpoAppGetCommand();