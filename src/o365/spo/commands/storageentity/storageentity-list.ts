import auth from '../../SpoAuth';
import Auth from '../../../../Auth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
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

interface Options extends GlobalOptions {
  appCatalogUrl: string;
}

interface TenantProperty {
  Comment?: string;
  Description?: string;
  Value: string;
}

class SpoStorageEntityListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_LIST}`;
  }

  public get description(): string {
    return 'Lists tenant properties stored on the specified SharePoint Online app catalog';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.appCatalogUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource} using refresh token ${auth.service.refreshToken}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): Promise<{ storageentitiesindex: string }> => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Loading all tenant properties...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving details for all tenant properties in ${args.options.appCatalogUrl}...`);
        }

        const requestOptions: any = {
          url: `${args.options.appCatalogUrl}/_api/web/AllProperties?$select=storageentitiesindex`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((web: { storageentitiesindex?: string }): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(web);
          cmd.log('');
        }

        try {
          if (!web.storageentitiesindex ||
            web.storageentitiesindex.trim().length === 0) {
            if (this.verbose) {
              cmd.log('No tenant properties found');
            }
            return;
          }

          const properties: { [key: string]: TenantProperty } = JSON.parse(web.storageentitiesindex);
          const keys: string[] = Object.keys(properties);
          if (keys.length === 0) {
            if (this.verbose) {
              cmd.log('No tenant properties found');
            }
          }
          else {
            keys.forEach((key: string): void => {
              const property: TenantProperty = properties[key];
              cmd.log(`Key:         ${key}`);
              cmd.log(`Value:       ${property.Value}`);
              cmd.log(`Description: ${(property.Description || 'not set')}`);
              cmd.log(`Comment:     ${(property.Comment || 'not set')}`);
              cmd.log('');
            });
          }
        }
        catch (e) {
          cmd.log(vorpal.chalk.red(`Error: ${e}`));
        }
        finally {
          cb();
        }
      }, (err: any): void => {
        cmd.log(vorpal.chalk.red(`Error: ${err}`));
        cb();
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-u, --appCatalogUrl <appCatalogUrl>',
      description: 'URL of the app catalog site'
    }];

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