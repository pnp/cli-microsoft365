import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { CommandOption, CommandCancel, CommandValidate } from '../../../../Command';
import { FormDigestInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import { SpoOperation } from '../site/SpoOperation';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  wait: boolean;
}

class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores the specified deleted site collection from tenant recycle bin';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.url = typeof args.options.url !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    this.dots = '';
    
    this.getSpoAdminUrl(cmd, this.debug)
    .then((adminUrl: string): Promise<FormDigestInfo> => {
      this.spoAdminUrl = adminUrl;

      return this.ensureFormDigest(this.spoAdminUrl, cmd, this.context, this.debug);
    })
    .then((res: FormDigestInfo): Promise<string> => {
      this.context = res;

      if (this.verbose) {
        cmd.log(`Restoring deleted site collection ${args.options.url}...`);
      }

      const requestOptions: any = {
        url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': this.context.FormDigestValue
        },
        body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">${args.options.url}</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      return request.post(requestOptions);
    })
    .then((res: string): Promise<void> => {
      return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          reject(response.ErrorInfo.ErrorMessage);
        }
        else {
          const operation: SpoOperation = json[json.length - 1];
          let isComplete: boolean = operation.IsComplete;
          if (!args.options.wait || isComplete) {
            resolve();
            return;
          }

          this.timeout = setTimeout(() => {
            this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, cmd, this.context as FormDigestInfo, this.dots, this.timeout);
          }, operation.PollingInterval);
        }
      });
    })
    .then((): void => {
      if (this.verbose) {
        cmd.log(vorpal.chalk.green('DONE'));
      }

      cb()
    }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public cancel(): CommandCancel {
    return (): void => {
      if (this.timeout) {
        clearTimeout(this.timeout);
      }
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'URL of the site to restore'
      },
      {
        option: '--wait',
        description: 'Wait for the deleted site to be restored'
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
    log(vorpal.find(commands.TENANT_RECYCLEBINITEM_RESTORE).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.

  Remarks:

    Restoring deleted site collections is by default asynchronous
    and depending on the current state of Office 365, might take up to few
    minutes. If you're building a script with steps that require the site to be
    fully restored, you should use the ${chalk.blue('--wait')} flag. When
    using this flag, the ${chalk.blue(this.getCommandName())} command will keep
    running until it received confirmation from Office 365 that the site
    has been fully restored.

  Examples:

    Restore a deleted site collection from tenant recycle bin
    ${commands.TENANT_RECYCLEBINITEM_RESTORE} --url https://contoso.sharepoint.com/sites/team

    Restore a deleted site collection from tenant recycle bin and wait for completion
    ${commands.TENANT_RECYCLEBINITEM_RESTORE} --url https://contoso.sharepoint.com/sites/team --wait
    `);
  }
}

module.exports = new SpoTenantRecycleBinItemRestoreCommand();