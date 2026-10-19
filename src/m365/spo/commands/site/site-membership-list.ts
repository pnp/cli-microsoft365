import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

enum Role {
  Owner = 'Owner',
  Member = 'Member',
  Visitor = 'Visitor'
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  role: zod.coercedEnum(Role).optional().alias('r')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
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
  public get name(): string {
    return commands.SITE_MEMBERSHIP_LIST;
  }

  public get description(): string {
    return `Retrieves information about default site groups' membership`;
  }

  public defaultProperties(): string[] | undefined {
    return ['email', 'name', 'userPrincipalName', 'associatedGroupType'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

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