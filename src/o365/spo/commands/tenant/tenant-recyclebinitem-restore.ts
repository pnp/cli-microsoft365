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
}

interface RestSpoOperation {
  HasTimedout: boolean;
  IsComplete: boolean;
  PollingInterval: number;
}

class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores the specified deleted Site Collection from Tenant Recycle Bin';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const wait: boolean = args.options.wait || false;
    let spoAdminUrl: string;

    this.getSpoAdminUrl(cmd, this.debug)
    .then((adminUrl: string): Promise<RestSpoOperation> => {
      spoAdminUrl = adminUrl;
      return this.restoreDeletedSite(args.options.url, spoAdminUrl, cmd);
    })
    .then((response: RestSpoOperation): void => {
      if (!response.HasTimedout && response.IsComplete) {
        if (this.verbose) {
          cmd.log('Site Collection restored');
        }

        cb();
      }
      else if (wait) {
        this.waitForRestoreDeletedSite(args.options.url, spoAdminUrl, cmd, response);
        cb();
      }
      else {
        cb('Site Collection has not been restored');
      }
    }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (this.timeout) {
        clearTimeout(this.timeout);
      }
    }
  }

  private waitForRestoreDeletedSite(siteToRestoreUrl: string, spoAdminUrl: string, cmd: CommandInstance, response: RestSpoOperation): void {
    this.timeout = setTimeout(() => {
      this.restoreDeletedSite(siteToRestoreUrl, spoAdminUrl, cmd)
      .then((respRetry: RestSpoOperation) => {
        this.waitForRestoreDeletedSite(siteToRestoreUrl, spoAdminUrl, cmd, respRetry);
      })
      .catch((err) => {
        cmd.log('Site Collection has not been restored');
        throw err;
      });
    }, response.PollingInterval);
  }

  private restoreDeletedSite(siteToRestoreUrl: string, spoAdminUrl: string, cmd: CommandInstance): Promise<RestSpoOperation> {
    if (this.verbose) {
      cmd.log('Restoring the deleted Site Collection');
    }

    if (this.debug) {
      cmd.log(`siteUrl : ${siteToRestoreUrl}`);
    }
    const requestOptions: any = {
      url: `${spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RestoreDeletedSite`,
      headers: {
        accept: 'application/json;odata=nometadata',
        contenttype: 'application/json;odata=nometadata',
      },
      body: {
        siteUrl: siteToRestoreUrl
      },
      json: true
    };

    return request.post(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the Site Collection to restore'
      },
      {
        option: '--wait',
        description: 'Wait for the Site Collection to be restored before completing the command'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.TENANT_RECYCLEBINITEM_RESTORE).helpInformation());
    log(
      ` ${chalk.yellow('Important:')} to use this command you have to have permissions to access
  the tenant admin site.
    
  Remarks:

    Restoring a Site Collection is by default asynchronous
    and depending on the current state of Office 365, might take up to few
    minutes. If you're building a script with steps that require the site to be
    fully restored, you should use the ${chalk.blue('--wait')} flag. When using this flag,
    the ${chalk.blue(this.getCommandName())} command will keep running until it received
    confirmation from Office 365 that the site has been fully restored.

  Examples:

    Restore a deleted Site Collection from Tenant Recycle Bin
    ${commands.TENANT_RECYCLEBINITEM_RESTORE} --url https://contoso.sharepoint.com/sites/team

    Restore a deleted Site Collection from Tenant Recycle Bin
    and wait for the restoring process to complete
    ${commands.TENANT_RECYCLEBINITEM_RESTORE} --url https://contoso.sharepoint.com/sites/team --wait
    `);
  }
}

module.exports = new SpoTenantRecycleBinItemRestoreCommand();