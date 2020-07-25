import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import request from '../../../../request';
import config from '../../../../config';
import commands from '../../commands';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { DeletedSitePropertiesEnumerable } from './DeletedSitePropertiesEnumerable';
import { CommandInstance } from '../../../../cli';

class SpoTenantRecycleBinItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Returns all modern and classic site collections in the tenant scoped recycle bin';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Url" ScalarProperty="true" /><Property Name="SiteId" ScalarProperty="true" /><Property Name="DaysRemaining" ScalarProperty="true" /><Property Name="Status" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="String">0</Parameter></Parameters></Method></ObjectPaths></Request>`
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

        const results: DeletedSitePropertiesEnumerable = json[json.length - 1];
        if (args.options.output === 'json') {
          cmd.log(results._Child_Items_);
        }
        else {
          cmd.log(results._Child_Items_.map((r: any) => {
            return {
              DaysRemaining: Number(r.DaysRemaining),
              DeletionTime: this.dateParser(r.DeletionTime as string),
              Url: r.Url
            };
          }).sort((a, b) => {
            const urlA = a.Url.toUpperCase();
            const urlB = b.Url.toUpperCase();
            if (urlA < urlB) {
              return -1;
            }
            if (urlA > urlB) {
              return 1;
            }
            return 0;
          }));
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }
  private dateParser(dateString: string): Date {
    const d: number[] = dateString.replace('/Date(', '').replace(')/', '').split(',').map(Number);
    return new Date(d[0], d[1], d[2], d[3], d[4], d[5], d[6]);
  }
}

module.exports = new SpoTenantRecycleBinItemListCommand();