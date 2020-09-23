import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  title: string;
  webPartData: string;
  addToQuickLaunch: boolean;
}

class SpoAppPageAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.APPPAGE_ADD}`;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.addToQuickLaunch = args.options.addToQuickLaunch;
    return telemetryProps;
  }

  public get description(): string {
    return 'Creates a single-part app page';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/sitepages/Pages/CreateFullPageApp`,
      headers: {
        'content-type': 'application/json;odata=nometadata',
        accept: 'application/json;odata=nometadata'
      },
      json: true,
      body: {
        title: args.options.title,
        addToQuickLaunch: args.options.addToQuickLaunch ? true : false,
        webPartDataAsJson: args.options.webPartData
      }
    };

    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);

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
        description: 'The URL of the site where the page should be created'
      },
      {
        option: '-t, --title <title>',
        description: 'The title of the page to be created'
      },
      {
        option: '-d, --webPartData <webPartData>',
        description: 'JSON string of the web part to put on the page'
      },
      {
        option: '--addToQuickLaunch',
        description: 'Set, to add the page to the quick launch'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    try {
      JSON.parse(args.options.webPartData);
    }
    catch (e) {
      return `Specified webPartData is not a valid JSON string. Error: ${e}`;
    }

    return true;
  }
}

module.exports = new SpoAppPageAddCommand();