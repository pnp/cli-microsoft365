import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { TermGroupCollection } from './TermGroupCollection';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: GlobalOptions;
}

class SpoTermGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.TERM_GROUP_LIST;
  }

  public get description(): string {
    return 'Lists taxonomy term groups';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Retrieving taxonomy term groups...`);
        }

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /><Property Id="9" ParentId="6" Name="Groups" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }

        const result: TermGroupCollection = json[json.length - 1];
        if (result._Child_Items_ && result._Child_Items_.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(result._Child_Items_.map(t => {
              t.CreatedDate = new Date(Number(t.CreatedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
              t.Id = t.Id.replace('/Guid(', '').replace(')/', '');
              t.LastModifiedDate = new Date(Number(t.LastModifiedDate.replace('/Date(', '').replace(')/', ''))).toISOString();
              return t;
            }));
          }
          else {
            cmd.log(result._Child_Items_.map(t => {
              return {
                Id: t.Id.replace('/Guid(', '').replace(')/', ''),
                Name: t.Name
              };
            }));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
}

module.exports = new SpoTermGroupListCommand();