import { Logger } from '../../../../cli/Logger.js';
import config from '../../../../config.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { FormDigestInfo, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { AdminResult, AdminUserResult, ISiteOwner, ISiteUser, ISPSite } from './SiteAdmin.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  primary?: boolean;
  asAdmin?: boolean;
}

class SpoSiteAdminAddCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ADMIN_ADD;
  }

  public get description(): string {
    return 'Adds a user or group as a site collection administrator';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        primary: !!args.options.primary,
        asAdmin: !!args.options.asAdmin
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
        option: '--primary'
      },
      {
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId &&
          !validation.isValidGuid(args.options.userId)) {
          return `'${args.options.userId}' is not a valid GUID for option 'userId'`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `'${args.options.userName}' is not a valid 'userName'`;
        }

        if (args.options.groupId &&
          !validation.isValidGuid(args.options.groupId)) {
          return `'${args.options.groupId}' is not a valid GUID for option 'groupId'`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['userId', 'userName', 'groupId', 'groupName'] });
  }

  #initTypes(): void {
    this.types.string.push('siteUrl', 'userId', 'userName', 'groupId', 'groupName');
    this.types.boolean.push('primary', 'asAdmin');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const loginNameToAdd = await this.getCorrectLoginName(args.options);
      if (args.options.asAdmin) {
        await this.callActionAsAdmin(logger, args, loginNameToAdd);
        return;
      }

      await this.callAction(logger, args, loginNameToAdd);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async callActionAsAdmin(logger: Logger, args: CommandArgs, loginNameToAdd: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Adding site administrator as an administrator...');
    }

    const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
    const siteId = await this.getSiteIdBasedOnUrl(args.options.siteUrl, logger);
    const siteAdmins = (await this.getSiteAdmins(adminUrl, siteId)).map(u => u.loginName);
    siteAdmins.push(loginNameToAdd);
    await this.setSiteAdminsAsAdmin(adminUrl, siteId, siteAdmins);
    if (args.options.primary) {
      await this.setPrimaryAdminAsAdmin(adminUrl, siteId, loginNameToAdd);
    }
  }

  private async getSiteIdBasedOnUrl(siteUrl: string, logger: Logger): Promise<string> {
    const siteGraphId = await spo.getSiteId(siteUrl, logger, this.verbose);
    const match = siteGraphId.match(/,([a-f0-9\-]{36}),/i);
    if (!match) {
      throw `Site with URL ${siteUrl} not found`;
    }

    return match[1];
  }

  private async getSiteAdmins(adminUrl: string, siteId: string): Promise<AdminUserResult[]> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`,
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
      const userPrincipalName = options.userName ? options.userName : await entraUser.getUpnByUserId(options.userId!);

      if (userPrincipalName) {
        return `i:0#.f|membership|${userPrincipalName}`;
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

  private async setSiteAdminsAsAdmin(adminUrl: string, siteId: string, admins: string[]): Promise<void> {
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
            admins
        }
      }
    };

    return request.post(requestOptions);
  }

  private async setPrimaryAdminAsAdmin(adminUrl: string, siteId: string, adminLogin: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      },
      data: {
        Owner: adminLogin,
        SetOwnerWithoutUpdatingSecondaryAdmin: true
      }
    };

    return request.patch(requestOptions);
  }

  private async callAction(logger: Logger, args: CommandArgs, loginNameToAdd: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Adding site administrator...');
    }

    const ensuredUserData = await this.ensureUser(args, loginNameToAdd);
    await this.setSiteAdmin(args.options.siteUrl, loginNameToAdd);

    if (args.options.primary) {
      const siteId = await this.getSiteId(args.options.siteUrl);
      const previousPrimaryOwner = await this.getSiteOwnerLoginName(args.options.siteUrl);
      await this.setPrimaryOwnerLoginFromSite(logger, args.options.siteUrl, siteId, ensuredUserData);
      await this.setSiteAdmin(args.options.siteUrl, previousPrimaryOwner);
    }
  }

  private async ensureUser(args: CommandArgs, loginName: string): Promise<ISiteUser> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.siteUrl}/_api/web/ensureuser`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      data: {
        logonName: loginName
      },
      responseType: 'json'
    };

    return request.post<ISiteUser>(requestOptions);
  }

  private async setSiteAdmin(siteUrl: string, loginName: string): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/web/siteusers('${formatting.encodeQueryParameter(loginName)}')`,
      headers: {
        'accept': 'application/json',
        'X-Http-Method': 'MERGE',
        'If-Match': '*'
      },
      data: { IsSiteAdmin: true },
      responseType: 'json'
    };
    return request.post(requestOptions);
  }

  private async getSiteId(siteUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/site?$select=Id`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<ISPSite>(requestOptions);
    return response.Id;
  }

  private async getSiteOwnerLoginName(siteUrl: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/site/owner?$select=LoginName`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.get<ISiteOwner>(requestOptions);
    return response.LoginName;
  }

  private async setPrimaryOwnerLoginFromSite(logger: Logger, siteUrl: string, siteId: string, loginName: ISiteUser): Promise<string> {
    const res: FormDigestInfo = await spo.ensureFormDigest(siteUrl, logger, undefined, this.debug);
    const body = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}"><Actions><SetProperty Id="10" ObjectPathId="2" Name="Owner"><Parameter ObjectPathId="3" /></SetProperty></Actions><ObjectPaths><Property Id="2" ParentId="0" Name="Site" /><Identity Id="3" Name="6d452ba1-40a8-8000-e00d-46e1adaa12bf|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${siteId}:u:${loginName.Id}" /><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`;

    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: {
        'X-RequestDigest': res.FormDigestValue
      },
      data: body
    };

    return request.post<string>(requestOptions);
  }
}

export default new SpoSiteAdminAddCommand();