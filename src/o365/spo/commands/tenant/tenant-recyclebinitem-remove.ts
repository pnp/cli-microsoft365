import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandCancel } from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  wait: boolean;
  confirm?: boolean;
}

interface RestSpoOperation {
  HasTimedout: boolean;
  IsComplete: boolean;
  PollingInterval: number;
}

class SpoTenantRecycleBinItemRemoveCommand extends SpoCommand {
  private timeout?: NodeJS.Timer;

  public get name (): string {
    return commands.TENANT_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified deleted Site Collection from Tenant Recycle Bin';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const wait: boolean = args.options.wait || false;
    let spoAdminUrl: string;

    const removeDeletedSite = (): void => {
      this.getSpoAdminUrl(cmd, this.debug)
      .then((adminUrl: string): Promise<RestSpoOperation> => {
        spoAdminUrl = adminUrl;
        return this.removeDeletedSite(args.options.url, spoAdminUrl, cmd);
      })
      .then((response: RestSpoOperation): void => {
        if (!response.HasTimedout && response.IsComplete) {
          if (this.verbose) {
            cmd.log('Site Collection removed');
          }

          cb();
        }
        else if (wait) {
          this.waitForRemoveDeletedSite(args.options.url, spoAdminUrl, cmd, response);
          cb();
        }
        else {
          cb('Site Collection has not been removed');
        }
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };    

    if (args.options.confirm) {
      removeDeletedSite();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the deleted Site Collection ${args.options.url} from Tenant Recycle Bin?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeDeletedSite();
        }
      });
    }
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (this.timeout) {
        clearTimeout(this.timeout);
      }
    }
  }

  private waitForRemoveDeletedSite(siteToRemoveUrl: string, spoAdminUrl: string, cmd: CommandInstance, response: RestSpoOperation): void {
    this.timeout = setTimeout(() => {
      this.removeDeletedSite(siteToRemoveUrl, spoAdminUrl, cmd)
      .then((respRetry: RestSpoOperation) => {
        this.waitForRemoveDeletedSite(siteToRemoveUrl, spoAdminUrl, cmd, respRetry);
      })
      .catch((err) => {
        cmd.log('Site Collection has not been removed');
        throw err;
      });
    }, response.PollingInterval);
  }

  private removeDeletedSite(siteToRemoveUrl: string, spoAdminUrl: string, cmd: CommandInstance): Promise<RestSpoOperation> {
    if (this.verbose) {
      cmd.log('Removing the deleted Site Collection');
    }

    if (this.debug) {
      cmd.log(`siteUrl : ${siteToRemoveUrl}`);
    }
    const requestOptions: any = {
      url: `${spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`,
      headers: {
        accept: 'application/json;odata=nometadata',
        contenttype: 'application/json;odata=nometadata',
      },
      body: {
        siteUrl: siteToRemoveUrl
      },
      json: true
    };

    return request.post(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the Site Collection to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the deleted Site Collection'
      },
      {
        option: '--wait',
        description: 'Wait for the Site Collection to be removed before completing the command'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.TENANT_RECYCLEBINITEM_REMOVE).helpInformation());
    log(
      ` ${chalk.yellow('Important:')} to use this command you have to have permissions to access
  the tenant admin site.
    
Remarks:

  Removing a Site Collection is by default asynchronous
  and depending on the current state of Office 365, might take up to few
  minutes. If you're building a script with steps that require the site to be
  fully removed, you should use the ${chalk.blue('--wait')} flag. When using this flag,
  the ${chalk.blue(this.getCommandName())} command will keep running until it received
  confirmation from Office 365 that the site has been fully removed.

Examples:

  Removes a deleted Site Collection from Tenant Recycle Bin
  ${commands.TENANT_RECYCLEBINITEM_REMOVE} --url https://contoso.sharepoint.com/sites/team

  Removes a deleted Site Collection from Tenant Recycle Bin
  and wait for the removing process to complete
  ${commands.TENANT_RECYCLEBINITEM_REMOVE} --url https://contoso.sharepoint.com/sites/team --wait
    `);
  }
}

module.exports = new SpoTenantRecycleBinItemRemoveCommand();