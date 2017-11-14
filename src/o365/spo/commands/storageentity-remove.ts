import auth from '../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../spo';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import VerboseOption from '../../../VerboseOption';
import Command, {
  CommandAction,
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../Command';
import appInsights from '../../../appInsights';
import Utils from '../../../Utils';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  appCatalogUrl: string;
  key: string;
  confirm?: boolean;
}

class SpoStorageEntityRemoveCommand extends Command {
  public get name(): string {
    return `${commands.STORAGEENTITY_REMOVE}`;
  }

  public get description(): string {
    return 'Removes tenant property stored on the specified SharePoint Online app catalog';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;

      appInsights.trackEvent({
        name: commands.STORAGEENTITY_REMOVE,
        properties: {
          verbose: verbose.toString(),
          confirm: (!(!args.options.confirm)).toString()
        }
      });

      if (!auth.site.connected) {
        this.log('Connect to a SharePoint Online tenant admin site first');
        cb();
        return;
      }

      if (!auth.site.isTenantAdminSite()) {
        this.log(`${auth.site.url} is not a tenant admin site. Connect to your tenant admin site and try again`);
        cb();
        return;
      }

      const removeTenantProperty = (): void => {
        this.log(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);

        auth
          .ensureAccessToken(auth.service.resource, this, verbose)
          .then((accessToken: string): Promise<ContextInfo> => {
            if (verbose) {
              this.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
            }

            const requestOptions: any = {
              url: `${auth.site.url}/_api/contextinfo`,
              headers: {
                authorization: `Bearer ${accessToken}`,
                accept: 'application/json;odata=nometadata'
              },
              json: true
            };

            if (verbose) {
              this.log('Executing web request...');
              this.log(requestOptions);
              this.log('');
            }

            return request.post(requestOptions);
          })
          .then((res: ContextInfo): Promise<string> => {
            if (verbose) {
              this.log('Response:');
              this.log(res);
              this.log('');
            }

            const requestOptions: any = {
              url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
              headers: {
                authorization: `Bearer ${auth.service.accessToken}`,
                'X-RequestDigest': res.FormDigestValue
              },
              body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
            };

            if (verbose) {
              this.log('Executing web request...');
              this.log(requestOptions);
              this.log('');
            }

            return request.post(requestOptions);
          })
          .then((res: string): void => {
            if (verbose) {
              this.log('Response:');
              this.log(res);
              this.log('');
            }

            const json: ClientSvcResponse = JSON.parse(res);
            const response: ClientSvcResponseContents = json[0];
            if (response.ErrorInfo) {
              this.log(vorpal.chalk.red(`Error: ${response.ErrorInfo.ErrorMessage}`));
            }
            else {
              this.log(vorpal.chalk.green('DONE'));
            }
            cb();
          }, (err: any): void => {
            this.log(vorpal.chalk.red(`Error: ${err}`));
            cb();
          });
      }

      if (args.options.confirm) {
        removeTenantProperty();
      }
      else {
        this.prompt({
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to delete the ${args.options.key} tenant property?`,
        }, (result: { continue: boolean }): void => {
          if (!result.continue) {
            cb();
          }
          else {
            removeTenantProperty();
          }
        });
      }
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --appCatalogUrl <appCatalogUrl>',
        description: 'URL of the app catalog site'
      },
      {
        option: '-k, --key <key>',
        description: 'Name of the tenant property to retrieve'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removal of a tenant property'
      }
    ];

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options && args.options.appCatalogUrl) {
        if (args.options.appCatalogUrl.indexOf('https://') !== 0 ||
          args.options.appCatalogUrl.indexOf('.sharepoint.com') === -1 ||
          args.options.appCatalogUrl.indexOf('/sites/') === -1) {
          return `${args.options.appCatalogUrl} is not a valid SharePoint Online app catalog URL`;
        }
        else {
          return true;
        }
      }
      else {
        return 'Missing required option appCatalogUrl';
      }
    };
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.STORAGEENTITY_REMOVE).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
                
  Remarks:

    To remove a tenant property, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Tenant properties are stored in the app catalog site associated with that tenant.
    To remove a property, you have to specify the absolute URL of the app catalog site.
    If you specify the URL of a site different than the app catalog, you will get an access denied error.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_REMOVE} -k AnalyticsId -u https://contoso.sharepoint.com/sites/appcatalog
      remove the ${chalk.grey('AnalyticsId')} tenant property. Yields a confirmation prompt before actually
      removing the property

    ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_REMOVE} -k AnalyticsId --confirm -u https://contoso.sharepoint.com/sites/appcatalog
      remove the ${chalk.grey('AnalyticsId')} tenant property. Suppresses the confirmation prompt

  More information:

    SharePoint Framework Tenant Properties
      https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
`);
    };
  }
}

module.exports = new SpoStorageEntityRemoveCommand();