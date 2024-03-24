import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';
import { SPOSitePropertiesEnumerable } from './SPOSitePropertiesEnumerable.js';

interface CommandArgs {
  options: Options;
}

interface AdminUserResult {
  email: string;
  loginName: string;
  name: string;
  userPrincipalName: string;
}

interface AdminResult {
  value: AdminUserResult[];
}

interface SiteUserResult {
  Email: string;
  Id: number;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
}

interface SiteResult {
  value: SiteUserResult[];
}

interface CommandResultItem {
  Id?: number;
  Email: string;
  IsPrimaryAdmin: boolean;
  LoginName: string;
  Title: string;
  PrincipalType?: number;
  PrincipalTypeString?: string;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  asAdmin?: boolean;
}

class SpoSiteAdminListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ADMIN_LIST;
  }

  public get description(): string {
    return 'Lists all administrators of a specific SharePoint site';
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
        asAdmin: typeof args.options.asAdmin !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (args.options.asAdmin) {
        await this.callActionAsAdmin(logger, args);
        return;
      }

      await this.callAction(logger, args);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async callActionAsAdmin(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await logger.logToStderr('Retrieving site administrators as an administrator...');
      const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const siteId = await this.getSiteIdBasedOnUrl(logger, args.options.siteUrl, adminUrl);

      const requestOptions: CliRequestOptions = {
        url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;charset=utf-8'
        }
      };

      const response: string = await request.post<string>(requestOptions);
      const responseContent: AdminResult = JSON.parse(response);
      const primaryAdminLoginName = await this.getPrimaryAdminLoginNameFromAdmin(adminUrl, siteId);

      const mappedResult = responseContent.value.map((u: AdminUserResult): CommandResultItem => {
        return {
          Id: undefined,
          Email: u.email,
          LoginName: u.loginName,
          Title: u.name,
          PrincipalType: undefined,
          PrincipalTypeString: undefined,
          IsPrimaryAdmin: u.loginName === primaryAdminLoginName
        };
      });
      await logger.log(mappedResult);
    }
    catch (err: any) {
      throw err;
    }
  }

  private async getSiteIdBasedOnUrl(logger: Logger, siteUrl: string, spoAdminUrl: string): Promise<string> {
    try {
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
    catch (error) {
      throw error;
    }
  }

  private async getPrimaryAdminLoginNameFromAdmin(adminUrl: string, siteId: string): Promise<string> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${adminUrl}/_api/SPO.Tenant/sites(%27${siteId}%27)?$select=OwnerLoginName`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;charset=utf-8'
        }
      };

      const response: string = await request.get<string>(requestOptions);
      const responseContent = JSON.parse(response);
      return responseContent.OwnerLoginName;
    }
    catch (err: any) {
      throw err;
    }
  }

  private async callAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await logger.logToStderr('Retrieving site administrators...');
      const requestOptions: CliRequestOptions = {
        url: `${args.options.siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin%20eq%20true`,
        method: 'GET',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const responseContent: SiteResult = await request.get<SiteResult>(requestOptions);
      const primaryOwnerLogin = await this.getPrimaryOwnerLoginFromSite(args.options.siteUrl);
      const mappedResult = responseContent.value.map((u: SiteUserResult): CommandResultItem => {
        return {
          Id: u.Id,
          LoginName: u.LoginName,
          Title: u.Title,
          PrincipalType: u.PrincipalType,
          PrincipalTypeString: ListPrincipalType[u.PrincipalType],
          Email: u.Email,
          IsPrimaryAdmin: u.LoginName === primaryOwnerLogin
        };
      });
      await logger.log(mappedResult);
    }
    catch (err: any) {
      throw err;
    }
  }

  private async getPrimaryOwnerLoginFromSite(siteUrl: string): Promise<string> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${siteUrl}/_api/site/owner`,
        method: 'GET',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const responseContent = await request.get<{ LoginName: string }>(requestOptions);
      return responseContent?.LoginName ?? undefined;
    }
    catch (err: any) {
      throw err;
    }
  }
}

export default new SpoSiteAdminListCommand();