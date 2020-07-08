import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on a SharePoint API that is currently
    in preview and is subject to change once the API reached general
    availability.

    By default, the hub site data is returned from the server's cache.
    To refresh the data with the latest updates, use the ${chalk.blue('-f, --forceRefresh')}
    option. Use this option, if you just made changes and need to see them right
    away.

    If the specified site is not connected to a hub site site and is not a hub
    site itself, no data will be retrieved.

  Examples:
  
    Get information about the hub site data for a site with URL ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${this.name} --webUrl https://contoso.sharepoint.com/sites/project-x

  More information:

    SharePoint hub sites new in Microsoft 365
      https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547
`);
  }
}

module.exports = new SpoHubSiteDataGetCommand();