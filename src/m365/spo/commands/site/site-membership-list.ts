import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
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
  associatedGroupType?: string;
}

class SpoSiteMembershipListCommand extends SpoCommand {
  public static readonly RoleNames: string[] = ['Owner', 'Member', 'Visitor'];

  public get name(): string {
    return commands.SITE_MEMBERSHIP_LIST;
  }

  public get description(): string {
    return `Retrieves information about default site groups' membership`;
  }

  public defaultProperties(): string[] | undefined {
    return ['email', 'name', 'userPrincipalName', 'associatedGroupType'];
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
        autocomplete: SpoSiteMembershipListCommand.RoleNames
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.role && !SpoSiteMembershipListCommand.RoleNames.some(roleName => roleName.toLocaleLowerCase() === args.options.role!.toLocaleLowerCase())) {
          return `'${args.options.role}' is not a valid value for option 'role'. Valid values are: ${SpoSiteMembershipListCommand.RoleNames.join(', ')}`;
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
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.verbose);
      const roleIds: string = this.getRoleIds(args.options.role);
      const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);

      const response = await odata.getAllItems<IMembershipResult>(`${spoAdminUrl}/_api/SPO.Tenant/sites/GetSiteUserGroups?siteId='${tenantSiteProperties.SiteId}'&userGroupIds=[${roleIds}]`);
      const result = args.options.output === 'json' ? this.mapResult(response, args.options.role) : this.mapListResult(response, args.options.role);

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRoleIds(role?: string): string {
    switch (role?.toLowerCase()) {
      case 'owner':
        return '0';
      case 'member':
        return '1';
      case 'visitor':
        return '2';
      default:
        return '0,1,2';
    }
  }

  private mapResult(response: IMembershipResult[], role?: string): IMembershipOutput {
    switch (role?.toLowerCase()) {
      case 'owner':
        return { AssociatedOwnerGroup: response[0].userGroup };
      case 'member':
        return { AssociatedMemberGroup: response[0].userGroup };
      case 'visitor':
        return { AssociatedVisitorGroup: response[0].userGroup };
      default:
        return {
          AssociatedOwnerGroup: response[0].userGroup,
          AssociatedMemberGroup: response[1].userGroup,
          AssociatedVisitorGroup: response[2].userGroup
        };
    }
  }

  private mapListResult(response: IMembershipResult[], role?: string): IUserInfo[] {
    const mapGroup = (groupIndex: number, groupType: string): IUserInfo[] =>
      response[groupIndex].userGroup.map(user => ({
        ...user,
        associatedGroupType: groupType
      }));

    switch (role?.toLowerCase()) {
      case 'owner':
        return mapGroup(0, 'Owner');
      case 'member':
        return mapGroup(0, 'Member');
      case 'visitor':
        return mapGroup(0, 'Visitor');
      default:
        return [
          ...mapGroup(0, 'Owner'),
          ...mapGroup(1, 'Member'),
          ...mapGroup(2, 'Visitor')
        ];
    }
  }
}

export default new SpoSiteMembershipListCommand();