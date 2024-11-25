import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { ListPrincipalType } from '../list/ListPrincipalType.js';
import { AdminResult, AdminUserResult, AdminCommandResultItem, SiteResult, SiteUserResult } from './SiteAdmin.js';

interface CommandArgs {
  options: Options;
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

  public defaultProperties(): string[] | undefined {
    return ['Id', 'LoginName', 'Title', 'PrincipalTypeString'];
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
    this.types.boolean.push('asAdmin');
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
    const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${tenantSiteProperties.SiteId}'`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const response = await request.post<AdminResult>(requestOptions);
    const primaryAdminLoginName = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, tenantSiteProperties.SiteId, logger, this.verbose);

    const mappedResult = response.value.map((u: AdminUserResult): AdminCommandResultItem => ({
      Email: u.email,
      LoginName: u.loginName,
      Title: u.name,
      IsPrimaryAdmin: u.loginName === primaryAdminLoginName
    }));
    await logger.log(mappedResult);
  }

  private async callAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Retrieving site administrators...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${args.options.siteUrl}/_api/web/siteusers?$filter=IsSiteAdmin eq true`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    const responseContent: SiteResult = await request.get<SiteResult>(requestOptions);
    const primaryOwnerLogin = await spo.getPrimaryOwnerLoginFromSite(args.options.siteUrl, logger, this.verbose);
    const mappedResult = responseContent.value.map((u: SiteUserResult): AdminCommandResultItem => ({
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
}

export default new SpoSiteAdminListCommand();