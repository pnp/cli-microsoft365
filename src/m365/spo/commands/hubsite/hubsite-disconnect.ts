import { ContextInfo } from '../../spo';
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
  url: string;
  confirm?: boolean;
}

class SpoHubSiteDisconnectCommand extends SpoCommand {
  public get name(): string {
    return `${commands.HUBSITE_DISCONNECT}`;
  }

  public get description(): string {
    return 'Disconnects the specifies site collection from its hub site';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const disconnectHubSite: () => void = (): void => {
      this
        .getRequestDigest(args.options.url)
        .then((res: ContextInfo): Promise<void> => {
          if (this.verbose) {
            cmd.log(`Disconnecting site collection ${args.options.url} from its hubsite...`);
          }

          const requestOptions: any = {
            url: `${args.options.url}/_api/site/JoinHubSite('00000000-0000-0000-0000-000000000000')`,
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

    if (args.options.confirm) {
      disconnectHubSite();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to disconnect the site collection ${args.options.url} from its hub site?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          disconnectHubSite();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site collection to disconnect from its hub site'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming disconnecting from the hub site'
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

  Examples:
  
    Disconnect the site collection with URL
    ${chalk.grey('https://contoso.sharepoint.com/sites/sales')} from its hub site. Will prompt
    for confirmation before disconnecting from the hub site.
      ${this.name} --url https://contoso.sharepoint.com/sites/sales

    Disconnect the site collection with URL
    ${chalk.grey('https://contoso.sharepoint.com/sites/sales')} from its hub site without
    prompting for confirmation
      ${this.name} --url https://contoso.sharepoint.com/sites/sales --confirm

  More information:

    SharePoint hub sites new in Microsoft 365
      https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547
`);
  }
}

module.exports = new SpoHubSiteDisconnectCommand();