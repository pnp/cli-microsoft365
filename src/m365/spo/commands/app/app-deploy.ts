import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

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

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    let appId: string = '';
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let appCatalogUrl: string = '';

    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<string> => {
        return this.getAppCatalogSiteUrl(cmd, spoUrl, args)
      })
      .then((_appCatalogUrl: string): Promise<{ UniqueId: string; }> => {
        appCatalogUrl = _appCatalogUrl;

        if (args.options.id) {
          if (this.verbose) {
            cmd.log(`Using the specified app id ${args.options.id}`);
          }

          return Promise.resolve({ UniqueId: args.options.id });
        }
        else {
          if (this.verbose) {
            cmd.log(`Looking up app id for app named ${args.options.name}...`);
          }

          const requestOptions: any = {
            url: `${appCatalogUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
            headers: {
              accept: 'application/json;odata=nometadata'
            },
            json: true
          };

          return request.get(requestOptions);
        }
      })
      .then((res: { UniqueId: string }): Promise<void> => {
        appId = res.UniqueId;

        if (this.verbose) {
          cmd.log(`Deploying app...`);
        }

        const requestOptions: any = {
          url: `${appCatalogUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${appId}')/deploy`,
          headers: {
            accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata;charset=utf-8'
          },
          body: { 'skipFeatureDeployment': args.options.skipFeatureDeployment || false },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the app to deploy. Specify the id or the name but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the app to deploy. Specify the id or the name but not both'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant or site collection app catalog. It must be specified when the scope is \'sitecollection\''
      },
      {
        option: '--skipFeatureDeployment',
        description: 'If the app supports tenant-wide deployment, deploy it to the whole tenant'
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

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
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
    };
  }
}

module.exports = new SpoAppDeployCommand();