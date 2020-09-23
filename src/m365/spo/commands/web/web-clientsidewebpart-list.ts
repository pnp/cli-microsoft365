import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { GetClientSideWebPartsRsp } from './GetClientSideWebPartsRsp';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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
          logger.log("No client-side web parts available for this site");
        }

        if (clientSideWebParts.length > 0) {
          logger.log(clientSideWebParts);
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
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

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.webUrl);
  }
}

module.exports = new SpoWebClientSideWebPartListCommand();