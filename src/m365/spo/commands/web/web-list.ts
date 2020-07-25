import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { WebPropertiesCollection } from "./WebPropertiesCollection";
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_LIST;
  }

  public get description(): string {
    return 'Lists subsites of the specified site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving all webs in site at ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web/webs`;

    if (args.options.output !== 'json') {
      requestUrl += '?$select=Title,Id,URL';
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get<WebPropertiesCollection>(requestOptions)
      .then((webProperties: WebPropertiesCollection): void => {
        if (args.options.output === 'json') {
          cmd.log(webProperties);
        }
        else {
          cmd.log(webProperties.value.map(l => {
            return {
              Title: l.Title,
              Url: l.Url,
              Id: l.Id
            };
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the parent site for which to retrieve the list of subsites'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoWebListCommand();