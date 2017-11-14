import auth from '../SpoAuth';
import Auth from '../../../Auth';
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

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  appCatalogUrl: string;
}

interface TenantProperty {
  Comment?: string;
  Description?: string;
  Value: string;
}

class SpoStorageEntityListCommand extends Command {
  public get name(): string {
    return `${commands.STORAGEENTITY_LIST}`;
  }

  public get description(): string {
    return 'Lists tenant properties stored on the specified SharePoint Online app catalog';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;

      appInsights.trackEvent({
        name: commands.STORAGEENTITY_LIST,
        properties: {
          verbose: verbose.toString()
        }
      });

      if (!auth.site.connected) {
        this.log('Connect to a SharePoint Online site first');
        cb();
        return;
      }

      const resource: string = Auth.getResourceFromUrl(args.options.appCatalogUrl);

      if (verbose) {
        this.log(`Retrieving access token for ${resource} using refresh token ${auth.service.refreshToken}...`);
      }

      auth
        .getAccessToken(resource, auth.service.refreshToken as string, this, verbose)
        .then((accessToken: string): Promise<{ storageentitiesindex: string }> => {
          if (verbose) {
            this.log(`Retrieved access token ${accessToken}. Loading all tenant properties...`);
          }

          this.log(`Retrieving details for all tenant properties in ${args.options.appCatalogUrl}...`);

          const requestOptions: any = {
            url: `${args.options.appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
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

          return request.get(requestOptions);
        })
        .then((web: { storageentitiesindex?: string }): void => {
          if (verbose) {
            this.log('Response:');
            this.log(web);
            this.log('');
          }

          try {
            if (!web.storageentitiesindex ||
              web.storageentitiesindex.trim().length === 0) {
              this.log('No tenant properties found');
              return;
            }

            const properties: { [key: string]: TenantProperty } = JSON.parse(web.storageentitiesindex);
            const keys: string[] = Object.keys(properties);
            if (keys.length === 0) {
              this.log('No tenant properties found');
            }
            else {
              keys.forEach((key: string): void => {
                const property: TenantProperty = properties[key];
                this.log(`Key:         ${key}`);
                this.log(`Value:       ${property.Value}`);
                this.log(`Description: ${(property.Description || 'not set')}`);
                this.log(`Comment:     ${(property.Comment || 'not set')}`);
                this.log('');
              });
            }
          }
          catch (e) {
            this.log(vorpal.chalk.red(`Error: ${e}`));
          }
          finally {
            cb();
          }
        }, (err: any): void => {
          this.log(vorpal.chalk.red(`Error: ${err}`));
          cb();
        });
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-u, --appCatalogUrl <appCatalogUrl>',
      description: 'URL of the app catalog site'
    }];

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
      log(vorpal.find(commands.STORAGEENTITY_LIST).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site using the
  ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To list tenant properties, you have to first connect to a SharePoint site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    Tenant properties are stored in the app catalog site. To list all tenant properties,
    you have to specify the absolute URL of the app catalog site. If you specify an incorrect
    URL, or the site at the given URL is not an app catalog site, no properties will be retrieved.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_LIST} -u https://contoso.sharepoint.com/sites/appcatalog
      list all tenant properties stored in the https://contoso.sharepoint.com/sites/appcatalog app catalog site

  More information:

    SharePoint Framework Tenant Properties
      https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
`);
    };
  }
}

module.exports = new SpoStorageEntityListCommand();