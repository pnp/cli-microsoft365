import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandCancel, CommandValidate } from '../../../../Command';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  wait?: boolean;
  confirm?: boolean;
}

interface RestSpoOperation {
  HasTimedout: boolean;
  IsComplete: boolean;
  PollingInterval: number;
}

class SpoTenantRecycleBinItemRemoveCommand extends SpoCommand {
  private timeout?: NodeJS.Timer;
  private readonly maxAttempts: number = 5;

  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified deleted site collection from tenant recycle bin';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.url = typeof args.options.url !== 'undefined';
    telemetryProps.wait = typeof args.options.wait !== 'undefined';
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const wait: boolean = args.options.wait || false;
    let spoAdminUrl: string;

    const removeDeletedSite = () => {
      this.getSpoAdminUrl(cmd, this.debug)
      .then((adminUrl: string): Promise<RestSpoOperation> => {
        spoAdminUrl = adminUrl;
        return this.removeDeletedSite(args.options.url, spoAdminUrl, cmd);
      })
      .then((response: RestSpoOperation): Promise<void> => {
        if (!response.HasTimedout && response.IsComplete) {
          if (this.verbose) {
            cmd.log('site collection removed');
          }

          return Promise.resolve();
        }
        else if (wait) {
          return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
            this.waitForRemoveDeletedSite(
              cmd,
              this,
              spoAdminUrl,
              args.options.url,
              resolve,
              reject,
              0
            );
          });
        }
        else {
          return Promise.reject('site collection has not been removed');
        }
      })
      .then(() => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb()
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
        message: `Are you sure you want to remove the deleted site collection ${args.options.url} from tenant recycle bin?`,
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

  private waitForRemoveDeletedSite(cmd: CommandInstance, command: SpoTenantRecycleBinItemRemoveCommand, spoAdminUrl: string, siteToRemoveUrl: string, resolve: () => void, reject: (error: any) => void, iteration: number): void {
    iteration++;

    new Promise((): Promise<void> => {
      return this.removeDeletedSite(siteToRemoveUrl, spoAdminUrl, cmd)
      .then((respRetry: RestSpoOperation) => {
        if (!respRetry.HasTimedout && respRetry.IsComplete) {
          if (this.verbose) {
            cmd.log('site collection removed');
          }

          resolve();
          return;
        }
        else if (respRetry.HasTimedout || iteration > this.maxAttempts) {
          reject('Operation timeout');
        }
        else {
          command.timeout = setTimeout(() => {
            command.waitForRemoveDeletedSite(cmd, command, spoAdminUrl, siteToRemoveUrl, resolve, reject, iteration);
          }, respRetry.PollingInterval);
        }
      })
      .catch((err) => {
        cmd.log('site collection has not been removed');
        reject(err);
      });
    });
  }

  private removeDeletedSite(siteToRemoveUrl: string, spoAdminUrl: string, cmd: CommandInstance): Promise<RestSpoOperation> {
    if (this.verbose) {
      cmd.log(`Removing site collection ${siteToRemoveUrl} from the recycle bin...`);
    }

    const requestOptions: any = {
      url: `${spoAdminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveDeletedSite`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
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
        description: 'URL of the site to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the deleted site collection'
      },
      {
        option: '--wait',
        description: 'Wait for the site collection to be removed before completing the command'
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

  public commandHelp(args: CommandArgs, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.TENANT_RECYCLEBINITEM_REMOVE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
    
  Remarks:

    Removing a site collection is by default asynchronous
    and depending on the current state of Office 365, might take up to few
    minutes. If you're building a script with steps that require the site to be
    fully removed, you should use the ${chalk.blue('--wait')} flag. When using this flag,
    the ${chalk.blue(commands.TENANT_RECYCLEBINITEM_REMOVE)} command will keep running until it received
    confirmation from Office 365 that the site has been fully removed.

  Examples:

    Removes a deleted site collection from tenant recycle bin
    ${commands.TENANT_RECYCLEBINITEM_REMOVE} --url https://contoso.sharepoint.com/sites/team

    Removes a deleted site collection from tenant recycle bin
    and wait for the removing process to complete
    ${commands.TENANT_RECYCLEBINITEM_REMOVE} --url https://contoso.sharepoint.com/sites/team --wait
    `);
  }
}

module.exports = new SpoTenantRecycleBinItemRemoveCommand();