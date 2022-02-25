import { Logger } from '../../../../cli';
import {
  CommandError
} from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { spo, ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { DeletedSitePropertiesEnumerable } from './DeletedSitePropertiesEnumerable';

class SpoTenantRecycleBinItemListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Returns all modern and classic site collections in the tenant scoped recycle bin';
  }

  public defaultProperties(): string[] | undefined {
    return ['DaysRemaining', 'DeletionTime', 'Url'];
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    spo
      .getSpoAdminUrl(logger, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;
        return spo.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          data: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties><Property Name="Url" ScalarProperty="true" /><Property Name="SiteId" ScalarProperty="true" /><Property Name="DaysRemaining" ScalarProperty="true" /><Property Name="Status" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="String">0</Parameter></Parameters></Method></ObjectPaths></Request>`
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
        if (args.options.output !== 'json') {
          results._Child_Items_.forEach(s => {
            s.DaysRemaining = Number(s.DaysRemaining);
            s.DeletionTime = this.dateParser(s.DeletionTime as string);
          });
        }
        logger.log(results._Child_Items_);

        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
  private dateParser(dateString: string): Date {
    const d: number[] = dateString.replace('/Date(', '').replace(')/', '').split(',').map(Number);
    return new Date(d[0], d[1], d[2], d[3], d[4], d[5], d[6]);
  }
}

module.exports = new SpoTenantRecycleBinItemListCommand();