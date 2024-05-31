import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { formatting } from '../../../../utils/formatting.js';
import { ClientSvcResponse, ClientSvcResponseContents, FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
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

interface Options extends GlobalOptions {
  siteUrl: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  asAdmin?: boolean;
  force?: boolean;
}

class SpoSiteAdminRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ADMIN_REMOVE;
  }

  public get description(): string {
    return 'Removes a user or group as site collection administrator';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        siteUrl: typeof args.options.siteUrl !== 'undefined',
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        force: typeof args.options.force !== 'undefined',
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
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '--groupId [groupId]'
      },
      {
        option: '--groupName [groupName]'
      },
      {
        option: '--asAdmin'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId &&
          !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid userName`;
        }

        if (args.options.groupId &&
          !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.force) {
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified user from the site administrators list ${args.options.siteUrl}?` });
        if (!result) {
          return;
        }
      }

      const loginNameToRemove = await this.getCorrectLoginName(args.options);
      if (args.options.asAdmin) {
        await this.callActionAsAdmin(logger, args, loginNameToRemove);
        return;
      }

      await this.callAction(logger, args, loginNameToRemove);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async callActionAsAdmin(logger: Logger, args: CommandArgs, loginNameToRemove: string): Promise<void> {
    await logger.logToStderr('Removing site administrator as an administrator...');

    const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
    const siteId = await this.getSiteIdBasedOnUrl(logger, args.options.siteUrl, adminUrl);
    const primaryAdminLoginName = await this.getPrimaryAdminLoginNameFromAdmin(adminUrl, siteId);

    if (loginNameToRemove === primaryAdminLoginName) {
      throw 'You cannot remove the primary site collection administrator.';
    }

    const existingAdmins = await this.getSiteAdmins(adminUrl, siteId);
    const adminsToSet = existingAdmins.filter(u => u.loginName.toLowerCase() !== loginNameToRemove.toLowerCase());
    await this.setSiteAdminsAsAdmin(adminUrl, siteId, adminsToSet);
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

  private async getPrimaryAdminLoginNameFromAdmin(adminUrl: string, siteId: string): Promise<string> {
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

  private async getSiteAdmins(adminUrl: string, siteId: string): Promise<AdminUserResult[]> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId=%27${siteId}%27`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      }
    };

    const response: string = await request.post<string>(requestOptions);
    const responseContent: AdminResult = JSON.parse(response);
    return responseContent.value;
  }

  private async getCorrectLoginName(options: Options): Promise<string> {
    if (options.userId || options.userName) {
      let requestUrl: string = `https://graph.microsoft.com/v1.0/users/${options.userId}`;

      if (options.userName) {
        requestUrl = `https://graph.microsoft.com/v1.0/users('${formatting.encodeQueryParameter(options.userName)}')`;
      }

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const user = await request.get<any>(requestOptions);
      if (user && user.userPrincipalName) {
        return `i:0#.f|membership|${user.userPrincipalName}`;
      }

      throw 'User not found.';
    }
    else {
      const group = options.groupId ? await entraGroup.getGroupById(options.groupId) : await entraGroup.getGroupByDisplayName(options.groupName!);
      //for entra groups, M365 groups have an associated email and security groups don't
      if (group?.mail) {
        //M365 group is prefixed with c:0o.c|federateddirectoryclaimprovider
        return `c:0o.c|federateddirectoryclaimprovider|${group.id}`;
      }
      else {
        //security group is prefixed with c:0t.c|tenant
        return `c:0t.c|tenant|${group?.id}`;
      }
    }
  }

  private async setSiteAdminsAsAdmin(adminUrl: string, siteId: string, admins: AdminUserResult[]): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPOInternalUseOnly.Tenant/SetSiteSecondaryAdministrators`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      },
      data: {
        secondaryAdministratorsFieldsData: {
          siteId: siteId,
          secondaryAdministratorLoginNames:
            admins.map(u => u.loginName)
        }
      }
    };

    await request.post<string>(requestOptions);
  }

  private async callAction(logger: Logger, args: CommandArgs, loginNameToRemove: string): Promise<void> {
    await logger.logToStderr('Removing site administrator...');

    const primaryOwnerLogin = await this.getPrimaryOwnerLoginFromSite(args.options.siteUrl);
    if (loginNameToRemove === primaryOwnerLogin) {
      throw 'You cannot remove the primary site collection administrator.';
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.siteUrl}/_api/web/siteusers('${formatting.encodeQueryParameter(loginNameToRemove)}')`,
      headers: {
        'accept': 'application/json',
        'X-Http-Method': 'MERGE',
        'If-Match': '*'
      },
      data: { IsSiteAdmin: false },
      responseType: 'json'
    };
    await request.post(requestOptions);
  }

  private async getPrimaryOwnerLoginFromSite(siteUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/site/owner`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const responseContent = await request.get<{ LoginName: string }>(requestOptions);
    return responseContent?.LoginName;
  }
}

export default new SpoSiteAdminRemoveCommand();