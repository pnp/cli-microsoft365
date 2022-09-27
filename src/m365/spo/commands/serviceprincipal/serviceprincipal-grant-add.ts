import { Logger } from '../../../../cli';
import config from '../../../../config';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { ClientSvcResponse, ClientSvcResponseContents, formatting, spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { SPOWebAppServicePrincipalPermissionGrant } from './SPOWebAppServicePrincipalPermissionGrant';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  resource: string;
  scope: string;
}

class SpoServicePrincipalGrantAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_GRANT_ADD;
  }

  public get description(): string {
    return 'Grants the service principal permission to the specified API';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GRANT_ADD];
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-r, --resource <resource>'
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        logger.logToStderr(`Retrieving request digest...`);
      }

      const resDigest = await spo.getRequestDigest(spoAdminUrl);

      const requestOptions: any = {
        url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: {
          'X-RequestDigest': resDigest.FormDigestValue
        },
        data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.resource)}</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.scope)}</Parameter></Parameters></Method></ObjectPaths></Request>`
      };

      const res = await request.post<string>(requestOptions);
      const json: ClientSvcResponse = JSON.parse(res);
      const response: ClientSvcResponseContents = json[0];
      if (response.ErrorInfo) {
        throw response.ErrorInfo.ErrorMessage;
      }
      else {
        const result: SPOWebAppServicePrincipalPermissionGrant = json[json.length - 1];
        delete result._ObjectType_;
        logger.log(result);
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

module.exports = new SpoServicePrincipalGrantAddCommand();