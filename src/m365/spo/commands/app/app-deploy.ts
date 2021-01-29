import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  id?: string;
  name?: string;
  scope?: string;
  skipFeatureDeployment?: boolean;
}

class SpoAppDeployCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_DEPLOY;
  }

  public get description(): string {
    return 'Deploys the specified app in the specified app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = (!(!args.options.id)).toString();
    telemetryProps.name = (!(!args.options.name)).toString();
    telemetryProps.appCatalogUrl = (!(!args.options.appCatalogUrl)).toString();
    telemetryProps.skipFeatureDeployment = args.options.skipFeatureDeployment || false;
    telemetryProps.scope = (!(!args.options.scope)).toString();
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    let appId: string = '';
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogUrl: string = '';

    this
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(logger, spoUrl, args)
      })
      .then((_appCatalogUrl: string): Promise<{ UniqueId: string; }> => {
        appCatalogUrl = _appCatalogUrl;

        if (args.options.id) {
          if (this.verbose) {
            logger.logToStderr(`Using the specified app id ${args.options.id}`);
          }

          return Promise.resolve({ UniqueId: args.options.id });
        }
        else {
          if (this.verbose) {
            logger.logToStderr(`Looking up app id for app named ${args.options.name}...`);
          }

          const requestOptions: any = {
            url: `${appCatalogUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            responseType: 'json'
          };

          return request.get(requestOptions);
        }
      })
      .then((res: { UniqueId: string }): Promise<void> => {
        appId = res.UniqueId;

        if (this.verbose) {
          logger.logToStderr(`Deploying app...`);
        }

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${appId}')/deploy`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata;charset=utf-8'
          },
          data: { 'skipFeatureDeployment': args.options.skipFeatureDeployment || false },
          responseType: 'json'
        };

        return request.post(requestOptions);
      })
      .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
        option: '--skipFeatureDeployment'
      },
      {
        option: '-s, --scope [scope]',
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

module.exports = new SpoAppDeployCommand();