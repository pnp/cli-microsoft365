import { Logger } from '../../../../cli/Logger';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { ClientSvcResponse, ClientSvcResponseContents, ContextInfo, spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  designPackageId?: string;
  url: string;
}

class SpoSiteCommSiteEnableCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_COMMSITE_ENABLE;
  }

  public get description(): string {
    return 'Enables communication site features on the specified site';
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
        designPackageId: typeof args.options.designPackageId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      },
      {
        option: '-i, --designPackageId [designPackageId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.designPackageId &&
          !validation.isValidGuid(args.options.designPackageId)) {
          return `${args.options.designPackageId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.url);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const designPackageId: string = args.options.designPackageId || '{d604dac3-50d3-405e-9ab9-d4713cda74ef}';

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);

      if (this.debug) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      const ctxInfo: ContextInfo = await spo.getRequestDigest(spoAdminUrl);
      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': ctxInfo.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.url)}</Parameter><Parameter Type="Guid">${formatting.escapeXml(designPackageId)}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
      };

      const res: string = await request.post(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
    } 
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteCommSiteEnableCommand();