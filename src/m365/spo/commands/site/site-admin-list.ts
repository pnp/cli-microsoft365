import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';

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
  Id: number | null;
  Email: string;
  IsPrimaryAdmin: boolean;
  LoginName: string;
  Title: string;
  PrincipalType: number | null;
  PrincipalTypeString: string | null;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        siteUrl: typeof args.options.siteUrl !== 'undefined',
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
        option: '--asAdmin'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('siteUrl');
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
    if (this.verbose) {
      await logger.logToStderr('Retrieving site administrators as an administrator...');
    }

    const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
    const siteId = await this.getSiteId(args.options.siteUrl, logger);
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      }
    };

    const response: string = await request.post<string>(requestOptions);
    const responseContent: AdminResult = JSON.parse(response);
    const primaryAdminLoginName = await this.getPrimaryAdminLoginNameFromAdmin(adminUrl, siteId);

    const mappedResult = responseContent.value.map((u: AdminUserResult): CommandResultItem => ({
      Id: null,
      Email: u.email,
      LoginName: u.loginName,
      Title: u.name,
      PrincipalType: null,
      PrincipalTypeString: null,
      IsPrimaryAdmin: u.loginName === primaryAdminLoginName
    }));
    await logger.log(mappedResult);
  }

  private async getSiteId(siteUrl: string, logger: Logger): Promise<string> {
    const siteGraphId = await spo.getSiteId(siteUrl, logger, this.verbose);
    const match = siteGraphId.match(/,([a-f0-9\-]{36}),/i);
    if (!match) {
      throw `Site with URL ${siteUrl} not found`;
    }

    return match[1];
  }

  private async getPrimaryAdminLoginNameFromAdmin(adminUrl: string, siteId: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      }
    };

    const response: string = await request.get<string>(requestOptions);
    const responseContent = JSON.parse(response);
    return responseContent.OwnerLoginName;
  }

  private async callAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving site administrators...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const responseContent: SiteResult = await request.get<SiteResult>(requestOptions);
    const primaryOwnerLogin = await this.getPrimaryOwnerLoginFromSite(args.options.siteUrl);
    const mappedResult = responseContent.value.map((u: SiteUserResult): CommandResultItem => ({
      Id: u.Id,
      LoginName: u.LoginName,
      Title: u.Title,
      PrincipalType: u.PrincipalType,
      PrincipalTypeString: ListPrincipalType[u.PrincipalType],
      Email: u.Email,
      IsPrimaryAdmin: u.LoginName === primaryOwnerLogin
    }));
    await logger.log(mappedResult);
  }

  private async getPrimaryOwnerLoginFromSite(siteUrl: string): Promise<string | null> {
    const requestOptions: CliRequestOptions = {
      url: `${siteUrl}/_api/site/owner`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const responseContent = await request.get<{ LoginName: string }>(requestOptions);
    return responseContent?.LoginName ?? null;
  }
}

export default new SpoSiteAdminListCommand();