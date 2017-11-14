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
  value: string;
  description?: string;
  comment?: string;
}

class SpoStorageEntitySetCommand extends Command {
  public get name(): string {
    return `${commands.STORAGEENTITY_SET}`;
  }

  public get description(): string {
    return 'Sets tenant property on the specified SharePoint Online app catalog';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;

      appInsights.trackEvent({
        name: commands.STORAGEENTITY_SET,
        properties: {
          verbose: verbose.toString()
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

      if (verbose) {
        this.log(`key option set. Retrieving access token for ${auth.service.resource}...`);
      }

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

          this.log(`Setting tenant property ${args.options.key} in ${args.options.appCatalogUrl}...`);

          const requestOptions: any = {
            url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              authorization: `Bearer ${auth.service.accessToken}`,
              'X-RequestDigest': res.FormDigestValue
            },
            body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="24" ObjectPathId="23" /><ObjectPath Id="26" ObjectPathId="25" /><ObjectPath Id="28" ObjectPathId="27" /><Method Name="SetStorageEntity" Id="29" ObjectPathId="27"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.key)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.value)}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.description || '')}</Parameter><Parameter Type="String">${Utils.escapeXml(args.options.comment || '')}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="23" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="25" ParentId="23" Name="GetSiteByUrl"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.appCatalogUrl)}</Parameter></Parameters></Method><Property Id="27" ParentId="25" Name="RootWeb" /></ObjectPaths></Request>`
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

            if (response.ErrorInfo.ErrorMessage.indexOf('Access denied.') > -1) {
              this.log('');
              this.log(`This error is often caused by invalid URL of the app catalog site. Verify, that the URL you specified as an argument of the ${commands.STORAGEENTITY_SET} command is a valid app catalog URL and try again.`);
              this.log('');
            }
          }
          else {
            this.log(vorpal.chalk.green('DONE'));
          }
          cb();
        }, (err: any): void => {
          this.log(vorpal.chalk.red(`Error: ${err}`));
          cb();
        });
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
        option: '-v, --value <value>',
        description: 'Value to set for the property'
      },
      {
        option: '-d, --description [description]',
        description: 'Description to set for the property'
      },
      {
        option: '-c, --comment [comment]',
        description: 'Comment to set for the property'
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
      log(vorpal.find(commands.STORAGEENTITY_SET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online tenant admin site,
  using the ${chalk.blue(commands.CONNECT)} command.
                
  Remarks:

    To set a tenant property, you have to first connect to a tenant admin site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso-admin.sharepoint.com`)}.
    If you are connected to a different site and will try to manage tenant properties,
    you will get an error.

    Tenant properties are stored in the app catalog site associated with that tenant.
    To set a property, you have to specify the absolute URL of the app catalog site.
    If you specify the URL of a site different than the app catalog, you will get an access denied error.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_SET} -k AnalyticsId -v 123 -d 'Web analytics ID' -c 'Use on all sites'
    -u https://contoso.sharepoint.com/sites/appcatalog
      set ${chalk.grey('123')} as the value of the ${chalk.grey('AnalyticsId')} tenant property. Also include a description
      and a comment for additional clarification of the usage of the property.

  More information:

    SharePoint Framework Tenant Properties
      https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
`);
    };
  }
}

module.exports = new SpoStorageEntitySetCommand();