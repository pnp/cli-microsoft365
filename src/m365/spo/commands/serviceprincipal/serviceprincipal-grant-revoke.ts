import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SPOWebAppServicePrincipalPermissionGrant } from './SPOWebAppServicePrincipalPermissionGrant.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  scope?: string;
}

class SpoServicePrincipalGrantRevokeCommand extends SpoCommand {
  public get name(): string {
    return commands.SERVICEPRINCIPAL_GRANT_REVOKE;
  }

  public get description(): string {
    return 'Revokes the specified set of permissions granted to the service principal';
  }

  public alias(): string[] | undefined {
    return [commands.SP_GRANT_REVOKE];
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-s, --scope [scope]'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl = await spo.getSpoAdminUrl(logger, this.debug);
      if (this.verbose) {
        await logger.logToStderr(`Retrieving request digest...`);
      }

      const reqDigest = await spo.getRequestDigest(spoAdminUrl);

      if (args.options.scope) {
        // revoke a single scope

        // #1 get the grant
        const getGrantRequestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="3" ParentId="1" Name="PermissionGrants" /><Method Id="5" ParentId="3" Name="GetByObjectId"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
        };
        const grantRequestRes = await request.post<string>(getGrantRequestOptions);
        const grantRequestJson: ClientSvcResponse = JSON.parse(grantRequestRes);
        const responseInfo: ClientSvcResponseContents = grantRequestJson[0];

        if (responseInfo.ErrorInfo) {
          throw responseInfo.ErrorInfo.ErrorMessage;
        }

        const grantInfo: SPOWebAppServicePrincipalPermissionGrant = grantRequestJson.pop();

        // #2 remove the scope from the grant
        const removeScopeRequestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="Remove" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">${grantInfo.ClientId}</Parameter><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">${formatting.escapeXml(args.options.scope)}</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="8" ParentId="1" Name="GrantManager" /><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`
        };
        const removeScopeRes = await request.post<string>(removeScopeRequestOptions);

        const removeScopeResJson: ClientSvcResponse = JSON.parse(removeScopeRes);
        const removeScopeResponse: ClientSvcResponseContents = removeScopeResJson[0];

        if (removeScopeResponse.ErrorInfo) {
          throw removeScopeResponse.ErrorInfo.ErrorMessage;
        }
      }
      else {
        // revoke the whole grant
        const requestOptions: CliRequestOptions = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': reqDigest.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">${formatting.escapeXml(args.options.id)}</Parameter></Parameters></Method></ObjectPaths></Request>`
        };

        const res = await request.post<string>(requestOptions);

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          throw response.ErrorInfo.ErrorMessage;
        }
      }
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }
}

export default new SpoServicePrincipalGrantRevokeCommand();