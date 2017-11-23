import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import {
  CommandHelp,
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  appCatalogUrl: string;
  key: string;
  confirm?: boolean;
}

class SpoStorageEntityRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_REMOVE}`;
  }

  public get description(): string {
    return 'Removes tenant property stored on the specified SharePoint Online app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return true;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeTenantProperty = (): void => {
      cmd.log(`Removing tenant property ${args.options.key} from ${args.options.appCatalogUrl}...`);

      auth
        .ensureAccessToken(auth.service.resource, cmd, this.verbose)
        .then((accessToken: string): Promise<ContextInfo> => {
          if (this.verbose) {
            cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
          }

          return this.getRequestDigest(cmd, this.verbose);
        })
        .then((res: ContextInfo): Promise<string> => {
          if (this.verbose) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`
          };

          if (this.verbose) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.post(requestOptions);
        })
        .then((res: string): void => {
          if (this.verbose) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          const json: ClientSvcResponse = JSON.parse(res);
          const response: ClientSvcResponseContents = json[0];
          if (response.ErrorInfo) {
            cmd.log(vorpal.chalk.red(`Error: ${response.ErrorInfo.ErrorMessage}`));
          }
          else {
            cmd.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => {
          cmd.log(vorpal.chalk.red(`Error: ${err}`));
          cb();
        });
    }

    if (args.options.confirm) {
      removeTenantProperty();
    }
    else {
      cmd.prompt({
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

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const result: boolean | string = SpoCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      if (result === false) {
        return 'Missing required option appCatalogUrl';
      }
      else {
        return result;
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