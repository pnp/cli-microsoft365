import * as chalk from 'chalk';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { NavigationNode } from './NavigationNode';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  location: string;
  webUrl: string;
}

class SpoNavigationNodeListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.NAVIGATION_NODE_LIST}`;
  }

  public get description(): string {
    return 'Lists nodes from the specified site navigation';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.location = args.options.location;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Retrieving navigation nodes...`);
    }

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/navigation/${args.options.location.toLowerCase()}`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<{ value: NavigationNode[] }>(requestOptions)
      .then((res: { value: NavigationNode[] }): void => {
        logger.log(res.value.map(n => {
          return {
            Id: n.Id,
            Title: n.Title,
            Url: n.Url
          };
        }));

        if (this.verbose) {
          logger.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'Absolute URL of the site for which to retrieve navigation'
      },
      {
        option: '-l, --location <location>',
        description: 'Navigation type to retrieve. Available options: QuickLaunch|TopNavigationBar',
        autocomplete: ['QuickLaunch', 'TopNavigationBar']
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
    if (isValidSharePointUrl !== true) {
      return isValidSharePointUrl;
    }

    if (args.options.location !== 'QuickLaunch' &&
      args.options.location !== 'TopNavigationBar') {
      return `${args.options.location} is not a valid value for the location option. Allowed values are QuickLaunch|TopNavigationBar`;
    }

    return true;
  }
}

module.exports = new SpoNavigationNodeListCommand();