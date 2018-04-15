import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { GetClientSideWebPartsRsp } from './GetClientSideWebPartsRsp';
import { Auth } from '../../../../Auth';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebClientSideWebPartListCommand extends SpoCommand {

  public get name(): string {
    return commands.WEB_CLIENTSIDEWEBPART_LIST;
  }

  public get description(): string {
    return 'Lists available client-side web parts';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetClientSideWebParts`,
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
      .then((res: GetClientSideWebPartsRsp): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        const clientSideWebParts: any[] = [];
        res.value.forEach(component => {
          if (component.ComponentType === 1) {
            clientSideWebParts.push({
              Id: component.Id.replace("{", "").replace("}", ""),
              Name: component.Name,
              Title: JSON.parse(component.Manifest).preconfiguredEntries[0].title.default
            });
          }
        });

        if (clientSideWebParts.length === 0 && this.verbose) {
          cmd.log("No client-side web parts available for this site");
        }

        if (clientSideWebParts.length > 0) {
          cmd.log(clientSideWebParts);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site for which to retrieve the information'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required option webUrl missing';
      }
      else {
        const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
      }
      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.

  Remarks:
  
    To get the list of available client-side web parts, you have to first
    connect to SharePoint using the ${chalk.blue(commands.CONNECT)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
  
  Examples:

    Lists all the available client-side web parts for the specified site
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com
    ` );
  }
}

module.exports = new SpoWebClientSideWebPartListCommand();