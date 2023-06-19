import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo, SpoOperation } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  wait?: boolean;
  confirm?: boolean;
}

class SpoTenantRecycleBinItemRemoveCommand extends SpoCommand {
  private context?: FormDigestInfo;
  private spoAdminUrl?: string;

  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified deleted site collection from tenant recycle bin';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        wait: typeof args.options.wait !== 'undefined',
        confirm: typeof args.options.confirm !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--wait'
      },
      {
        option: '--confirm'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.confirm) {
      await this.removeDeletedSite(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the deleted site collection ${args.options.siteUrl} from tenant recycle bin?`
      });

      if (result.continue) {
        await this.removeDeletedSite(logger, args);
      }
    }
  }

  private async removeDeletedSite(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      const res: FormDigestInfo = await spo.ensureFormDigest(this.spoAdminUrl, logger, this.context, this.debug);
      if (this.verbose) {
        logger.logToStderr(`Removing deleted site collection ${args.options.siteUrl}...`);
      }

      const requestOptions: any = {
        url: `${this.spoAdminUrl as string}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': res.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="15" ParentId="1" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.siteUrl)}</Parameter></Parameters></Method><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const processQuery: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(processQuery);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const operation: SpoOperation = json[json.length - 1];
        const isComplete: boolean = operation.IsComplete;
        if (!args.options.wait || isComplete) {
          return;
        }

        await new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
          setTimeout(() => {
            spo.waitUntilFinished({
              operationId: JSON.stringify(operation._ObjectIdentity_),
              siteUrl: this.spoAdminUrl as string,
              resolve,
              reject,
              logger,
              currentContext: this.context as FormDigestInfo,
              debug: this.debug,
              verbose: this.verbose
            });
          }, operation.PollingInterval);
        });
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoTenantRecycleBinItemRemoveCommand();