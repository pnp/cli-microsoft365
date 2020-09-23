import * as chalk from 'chalk';
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
    return `${commands.APPPAGE_SET}`;
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
      json: true,
      body: {
        serverRelativeUrl: `${Utils.getServerRelativePath(args.options.webUrl, '')}/SitePages/${args.options.pageName}`,
        webPartDataAsJson: args.options.webPartData
      }
    };

    request.post(requestOptions).then((res: any): void => {
      if (this.verbose) {
        logger.log(chalk.green('DONE'));
      }
      cb();
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the page to update is located'
      },
      {
        option: '-n, --pageName <pageName>',
        description: 'The name of the page to be updated, eg. page.aspx'
      },
      {
        option: '-d, --webPartData <webPartData>',
        description: 'JSON string of the web part to update on the page'
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