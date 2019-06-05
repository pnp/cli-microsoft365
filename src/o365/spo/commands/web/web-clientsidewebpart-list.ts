import commands from '../../commands';
import request from '../../../../request';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { GetClientSideWebPartsRsp } from './GetClientSideWebPartsRsp';
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
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/GetClientSideWebParts`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<GetClientSideWebPartsRsp>(requestOptions)
      .then((res: GetClientSideWebPartsRsp): void => {
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
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Lists all the available client-side web parts for the specified site
      ${this.name} --webUrl https://contoso.sharepoint.com
    ` );
  }
}

module.exports = new SpoWebClientSideWebPartListCommand();