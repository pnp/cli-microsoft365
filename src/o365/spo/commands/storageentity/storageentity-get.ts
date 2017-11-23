import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import VerboseOption from '../../../../VerboseOption';
import {
  CommandHelp,
  CommandOption
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

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

class SpoStorageEntityGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_GET}`;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving access token for ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.verbose)
      .then((accessToken: string): Promise<TenantProperty> => {
        if (this.verbose) {
          cmd.log(`Retrieved access token ${accessToken}. Loading details for the ${args.options.key} tenant property...`);
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/web/GetStorageEntity('${encodeURIComponent(args.options.key)}')`,
          headers: {
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        if (this.verbose) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((property: TenantProperty): void => {
        if (this.verbose) {
          cmd.log('Property:');
          cmd.log(property);
          cmd.log('');
        }

        if (property["odata.null"] === true) {
          cmd.log(`Property with key ${args.options.key} not found`);
        }
        else {
          cmd.log(`Details for tenant property ${args.options.key}:`);
          cmd.log(`  Value:       ${property.Value}`);
          cmd.log(`  Description: ${(property.Description || 'not set')}`);
          cmd.log(`  Comment:    ${(property.Comment || 'not set')}`);
        }
        cmd.log('');
        cb();
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-k, --key <key>',
      description: 'Name of the tenant property to retrieve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
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