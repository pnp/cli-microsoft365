import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SPOSitePropertiesEnumerable } from '../site/SPOSitePropertiesEnumerable.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  role?: string;
}

interface IMembershipResult {
  userGroup: IUserInfo[];
}

interface IMembershipOutput {
  AssociatedOwnerGroup?: IUserInfo[];
  AssociatedMemberGroup?: IUserInfo[];
  AssociatedVisitorGroup?: IUserInfo[];
}

interface IUserInfo {
  email: string;
  loginName: string;
  name: string;
  userPrincipalName: string;
}

class SpoTenantSiteMemberShipListCommand extends SpoCommand {
  public static readonly RoleName: string[] = ['Owner', 'Member', 'Visitor'];

  public get name(): string {
    return commands.TENANT_SITE_MEMBERSHIP_LIST;
  }

  public get description(): string {
    return `Retrieve information about default site groups' membership.`;
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
        siteUrl: typeof args.options.siteUrl !== 'undefined',
        role: typeof args.options.role !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-r, --role [role]',
        autocomplete: SpoTenantSiteMemberShipListCommand.RoleName
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.role && SpoTenantSiteMemberShipListCommand.RoleName.indexOf(args.options.role) === -1) {
          return 'The value of parameter role must be Visitor|Member|Owner';
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const roleIds: string = this.getRoleIds(args.options.role);
      const siteId = await this.getSiteIdBasedOnUrl(logger, args.options.siteUrl, spoAdminUrl);

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId=%27${siteId}%27&userGroupIds=[${roleIds}]`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-Type': 'application/json'
        },
        responseType: 'json'
      };

      const response = await request.get<{ value: any }>(requestOptions);
      const result = this.mapResult(response.value, args);

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRoleIds(role: string | undefined): string {
    if (role === 'Owner') {
      return '0';
    }
    else if (role === 'Member') {
      return '1';
    }
    else if (role === 'Visitor') {
      return '2';
    }
    else {
      return '0,1,2';
    }
  }

  private async getSiteIdBasedOnUrl(logger: Logger, siteUrl: string, spoAdminUrl: string): Promise<string> {
    const res: FormDigestInfo = await spo.ensureFormDigest(spoAdminUrl, logger, undefined, this.debug);

    const urlFilter = formatting.escapeXml(`Url -eq '${siteUrl}'`);
    const requestBody: string = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="SiteId" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="GetSitePropertiesFromSharePointByFilters"><Parameters><Parameter TypeId="{b92aeee2-c92c-4b67-abcc-024e471bc140}"><Property Name="Filter" Type="String">${urlFilter}</Property><Property Name="IncludeDetail" Type="Boolean">false</Property><Property Name="IncludePersonalSite" Type="Enum">1</Property><Property Name="StartIndex" Type="String">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;

    const requestOptions: CliRequestOptions = {
      url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': res.FormDigestValue
      },
      data: requestBody
    };

    const response: string = await request.post<string>(requestOptions);
    const jsonData: ClientSvcResponse = JSON.parse(response);
    const responseContent: ClientSvcResponseContents = jsonData[0];

    if (responseContent.ErrorInfo) {
      throw responseContent.ErrorInfo.ErrorMessage;
    }

    const sites: SPOSitePropertiesEnumerable = jsonData[jsonData.length - 1];
    const siteId = sites?._Child_Items_?.[0]?.SiteId ?? undefined;

    if (!siteId) {
      throw 'Failed to obtain site Id from the provided site URL.';
    }

    const guid = siteId.replace('/Guid(', '').replace(')/', '');

    if (!validation.isValidGuid(guid)) {
      throw 'Failed to obtain site Id from the provided site URL.';
    }
    return guid;
  }

  private mapResult(response: IMembershipResult[], args: CommandArgs): IMembershipOutput {
    switch (args.options.role) {
      case 'Owner':
        return { 'AssociatedOwnerGroup': response[0].userGroup };
      case 'Member':
        return { 'AssociatedMemberGroup': response[0].userGroup };
      case 'Visitor':
        return { 'AssociatedVisitorGroup': response[0].userGroup };
      default:
        return {
          'AssociatedOwnerGroup': response[0].userGroup,
          'AssociatedMemberGroup': response[1].userGroup,
          'AssociatedVisitorGroup': response[2].userGroup
        };
    }
  }
}

export default new SpoTenantSiteMemberShipListCommand();