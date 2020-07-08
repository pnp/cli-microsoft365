import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  hubSiteId: string;
}

class SpoHubSiteConnectCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_CONNECT}`;
  }

  public get description(): string {
    return 'Connects the specified site collection to the given hub site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    this
      .getRequestDigest(args.options.url)
      .then((res: ContextInfo): Promise<void> => {
        const requestOptions: any = {
          url: `${args.options.url}/_api/site/JoinHubSite('${encodeURIComponent(args.options.hubSiteId)}')`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The URL of the site collection to connect to the hub site'
      },
      {
        option: '-i, --hubSiteId <hubSiteId>',
        description: 'The ID of the hub site to which to connect the site collection'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.hubSiteId) {
        return 'Required parameter hubSiteId missing';
      }

      if (!Utils.isValidGuid(args.options.hubSiteId)) {
        return `${args.options.hubSiteId} is not a valid GUID`;
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

    If the specified site collection is already connected to a hub site,
    it will be disconnected and connected to the newly specified hub site.

    If the specified ${chalk.grey('hubSiteId')} doesn't point to a valid hub
    site, you will get a ${chalk.grey('ResourceNotFoundException')} error.

  Examples:
  
    Connect the site collection with URL
    ${chalk.grey('https://contoso.sharepoint.com/sites/contoso-sales')} to the hub site
    with ID ${chalk.grey('255a50b2-527f-4413-8485-57f4c17a24d1')}
      ${this.name} --url https://contoso.sharepoint.com/sites/contoso-sales --hubSiteId 255a50b2-527f-4413-8485-57f4c17a24d1

  More information:

    SharePoint hub sites new in Microsoft 365
      https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547
`);
  }
}

module.exports = new SpoHubSiteConnectCommand();