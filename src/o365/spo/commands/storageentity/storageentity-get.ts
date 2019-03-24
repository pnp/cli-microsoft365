import auth from '../../SpoAuth';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { TenantProperty } from './TenantProperty';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  key: string;
}

class SpoStorageEntityGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.STORAGEENTITY_GET}`;
  }

  public get description(): string {
    return 'Get details for the specified tenant property';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.debug) {
      cmd.log(`Retrieving access token for ${auth.service.resource}...`);
    }

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): Promise<TenantProperty> => {
        if (this.debug) {
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

        return request.get(requestOptions);
      })
      .then((property: TenantProperty): void => {
        if (property["odata.null"] === true) {
          if (this.verbose) {
            cmd.log(`Property with key ${args.options.key} not found`);
          }
        }
        else {
          cmd.log({
            Key: args.options.key,
            Value: property.Value,
            Description: property.Description,
            Comment: property.Comment
          });
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-k, --key <key>',
      description: 'Name of the tenant property to retrieve'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.STORAGEENTITY_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site using the
        ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To get details of a tenant property, you have to first log in to a SharePoint site using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    Tenant properties are stored in the app catalog site associated with the site to which you are
    currently logged in. When retrieving the specified tenant property, SharePoint will automatically
    find the associated app catalog and try to retrieve the property from it.

  Examples:
  
    Show the value, description and comment of the ${chalk.grey('AnalyticsId')} tenant property
      ${chalk.grey(config.delimiter)} ${commands.STORAGEENTITY_GET} -k AnalyticsId

  More information:

    SharePoint Framework Tenant Properties
      https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
`);
  }
}

module.exports = new SpoStorageEntityGetCommand();