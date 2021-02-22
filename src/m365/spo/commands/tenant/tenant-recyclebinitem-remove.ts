import { Cli, Logger } from '../../../../cli';
import { CommandOption } from '../../../../Command';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo } from '../../spo';
import { SpoOperation } from '../site/SpoOperation';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
  wait?: boolean;
  confirm?: boolean;
}

class SpoTenantRecycleBinItemRemoveCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;
  private dots?: string;
  private timeout?: NodeJS.Timer;

  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified deleted site collection from tenant recycle bin';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.wait = typeof args.options.wait !== 'undefined';
    telemetryProps.confirm = typeof args.options.confirm !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    const removeDeletedSite = () => {
      this
        .getSpoAdminUrl(logger, this.debug)
        .then((adminUrl: string): Promise<FormDigestInfo> => {
          this.spoAdminUrl = adminUrl;

          return this.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
        })
        .then((res: FormDigestInfo): Promise<string> => {
          this.context = res;

          if (this.verbose) {
            logger.logToStderr(`Removing deleted site collection ${args.options.url}...`);
          }

          const requestOptions: any = {
            url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
            headers: {
              'X-RequestDigest': this.context.FormDigestValue
            },
            data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${Utils.escapeXml(args.options.url)}</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
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
              const isComplete: boolean = operation.IsComplete;
              if (!args.options.wait || isComplete) {
                resolve();
                return;
              }

              this.timeout = setTimeout(() => {
                this.waitUntilFinished(JSON.stringify(operation._ObjectIdentity_), this.spoAdminUrl as string, resolve, reject, logger, this.context as FormDigestInfo, this.dots, this.timeout);
              }, operation.PollingInterval);
            }
          });
        })
        .then(_ => cb(), (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
    };

    if (args.options.confirm) {
      removeDeletedSite();
    }
    else {
      Cli.prompt({
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>'
      },
      {
        option: '--wait'
      },
      {
        option: '--confirm'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    return SpoCommand.isValidSharePointUrl(args.options.url);
  }
}

module.exports = new SpoTenantRecycleBinItemRemoveCommand();