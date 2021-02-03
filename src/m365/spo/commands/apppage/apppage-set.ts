import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  pageName: string;
  webPartData: string;
}

class SpoAppPageSetCommand extends SpoCommand {
  public get name(): string {
    return commands.APPPAGE_SET;
  }

  public get description(): string {
    return 'Updates the single-part app page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/UpdateFullPageApp`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        serverRelativeUrl: `${Utils.getServerRelativePath(args.options.webUrl, '')}/SitePages/${args.options.pageName}`,
        webPartDataAsJson: args.options.webPartData
      }
    };

    request
      .post(requestOptions)
      .then(_ => cb(),
        (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --pageName <pageName>'
      },
      {
        option: '-d, --webPartData <webPartData>'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.webUrl) {
      return 'Required parameter webUrl missing';
    }
    if (!args.options.pageName) {
      return 'Required parameter pageName missing';
    }
    if (!args.options.webPartData) {
      return 'Required parameter webPartData missing';
    }
    try {
      JSON.parse(args.options.webPartData);
    } catch (e) {
      return `Specified webPartData is not a valid JSON string. Error: ${e}`;
    }
    return true;
  }
}
module.exports = new SpoAppPageSetCommand();