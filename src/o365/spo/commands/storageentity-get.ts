import auth from '../SpoAuth';
import config from '../../../config';
import * as request from 'request-promise-native';
import commands from '../commands';
import VerboseOption from '../../../VerboseOption';
import Command, {
  CommandAction,
  CommandHelp,
  CommandOption
} from '../../../Command';
import appInsights from '../../../appInsights';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends VerboseOption {
  key: string;
}

interface TenantProperty {
  "odata.null": boolean,
  Comment?: string;
  Description?: string;
  Value: string;
}

class SpoStorageEntityGetCommand extends Command {
  public get name(): string {
    return `${commands.STORAGEENTITY_GET}`;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
  }

  public get action(): CommandAction {
    return function (args: CommandArgs, cb: () => void) {
      const verbose: boolean = args.options.verbose || false;

      appInsights.trackEvent({
        name: commands.STORAGEENTITY_GET,
        properties: {
          verbose: verbose.toString()
        }
      });

      if (!auth.site.connected) {
        this.log('Connect to a SharePoint Online site first');
        cb();
        return;
      }

      if (verbose) {
        this.log(`key option set. Retrieving access token for ${auth.service.resource}...`);
      }

      auth
        .ensureAccessToken(auth.service.resource, this, verbose)
        .then((accessToken: string): Promise<TenantProperty> => {
          if (verbose) {
            this.log(`Retrieved access token ${accessToken}. Loading details for the ${args.options.key} tenant property...`);
          }

          const requestOptions: any = {
            url: `${auth.site.url}/_api/web/GetStorageEntity('${encodeURIComponent(args.options.key)}')`,
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
        .then((property: TenantProperty): void => {
          if (verbose) {
            this.log('Property:');
            this.log(property);
            this.log('');
          }

          if (property["odata.null"] === true) {
            this.log(`Property with key ${args.options.key} not found`);
          }
          else {
            this.log(`Details for tenant property ${args.options.key}:`);
            this.log(`  Value:       ${property.Value}`);
            this.log(`  Description: ${(property.Description || 'not set')}`);
            this.log(`  Comment:    ${(property.Comment || 'not set')}`);
          }
          this.log('');
          cb();
        }, (err: any): void => {
          this.log(vorpal.chalk.red(`Error: ${err}`));
          cb();
        });
    };
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-k, --key <key>',
      description: 'Name of the tenant property to retrieve'
    }];

    const parentOptions: CommandOption[] | undefined = super.options();
    if (parentOptions) {
      return options.concat(parentOptions);
    }
    else {
      return options;
    }
  }

  public help(): CommandHelp {
    return function (args: CommandArgs, log: (help: string) => void): void {
      const chalk = vorpal.chalk;
      log(vorpal.find(commands.STORAGEENTITY_GET).helpInformation());
      log(
        `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site using the
        ${chalk.blue(commands.CONNECT)} command.
        
  Remarks:

    To get details of a tenant property, you have to first connect to a SharePoint site using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    Tenant properties are stored in the app catalog site associated with the site to which you are
    currently connected. When retrieving the specified tenant property, SharePoint will automatically
    find the associated app catalog and try to retrieve the property from it.

  Examples:
  
    ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_GET} -k AnalyticsId
      show the value, description and comment of the ${chalk.grey('AnalyticsId')} tenant property

  More information:

    SharePoint Framework Tenant Properties
      https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
`);
    };
  }
}

module.exports = new SpoStorageEntityGetCommand();