import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  forceRefresh?: boolean;
}

class SpoHubSiteDataGetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_DATA_GET}`;
  }

  public get description(): string {
    return 'Get hub site data for the specified site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.forceRefresh = args.options.forceRefresh === true;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log('Retrieving hub site data...');
    }

    const forceRefresh: boolean = args.options.forceRefresh === true;

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/HubSiteData(${forceRefresh})`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (res['odata.null'] !== true) {
          cmd.log(JSON.parse(res.value));
        }
        else {
          if (this.verbose) {
            cmd.log(`${args.options.webUrl} is not connected to a hub site and is not a hub site itself`);
          }
        }

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site for which to retrieve hub site data'
      },
      {
        option: '-f, --forceRefresh',
        description: `Set, to refresh the server cache with the latest updates`
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

module.exports = new SpoHubSiteDataGetCommand();