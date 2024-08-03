import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

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
    return `Retrieve information about default site groups' membership`;
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

  #initTypes(): void {
    this.types.string.push('role', 'siteUrl');
  };

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const roleIds: string = this.getRoleIds(args.options.role);
      const siteId = await this.getSiteId(args.options.siteUrl, logger);

      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${siteId}'&userGroupIds=[${roleIds}]`,
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

  private async getSiteId(siteUrl: string, logger: Logger): Promise<string> {
    const siteGraphId = await spo.getSiteId(siteUrl, logger, this.verbose);
    const match = siteGraphId.match(/,([a-f0-9\-]{36}),/i);
    if (!match) {
      throw `Site with URL ${siteUrl} not found`;
    }

    return match[1];
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