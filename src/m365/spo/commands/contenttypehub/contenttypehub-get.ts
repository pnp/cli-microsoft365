import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';  
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
}

class SpoContentTypeHubGetCommand extends SpoCommand {
  public get name(): string {
    return commands.CONTENTTYPEHUB_GET;
  }

  public get description(): string {
    return 'Returns the URL of the SharePoint Content Type Hub of the Tenant';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoUrl: string = '';

    this
      .getSpoUrl(cmd,this.debug)
      .then((_spoUrl: string): Promise<ContextInfo> => {
        spoUrl = _spoUrl;
        return this.getRequestDigest(spoUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Retrieving Content Type Hub URL`);
        }

        const requestOptions: any = {
          url: `${spoUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}">
  <Actions>
    <ObjectPath Id="2" ObjectPathId="1" />
    <ObjectIdentityQuery Id="3" ObjectPathId="1" />
    <ObjectPath Id="5" ObjectPathId="4" />
    <ObjectIdentityQuery Id="6" ObjectPathId="4" />
    <Query Id="7" ObjectPathId="4">
      <Query SelectAllProperties="false">
        <Properties>
          <Property Name="ContentTypePublishingHub" ScalarProperty="true" />
        </Properties>
      </Query>
    </Query>
  </Actions>
  <ObjectPaths>
    <StaticMethod Id="1" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" />
    <Method Id="4" ParentId="1" Name="GetDefaultSiteCollectionTermStore" />
  </ObjectPaths>
</Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
        }
        else {
          const result: any = {
            ContentTypePublishingHub: json[json.length - 1]["ContentTypePublishingHub"]
          } 
          cmd.log(result);
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
}

module.exports = new SpoContentTypeHubGetCommand();