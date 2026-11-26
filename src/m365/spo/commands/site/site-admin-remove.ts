import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { AdminResult, AdminUserResult } from './SiteAdmin.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  userId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  userName: z.string()
    .refine(name => validation.isValidUserPrincipalName(name), {
      error: e => `${e.input} is not a valid userName`
    })
    .optional(),
  groupId: z.string()
    .refine(id => validation.isValidGuid(id), {
      error: e => `${e.input} is not a valid GUID`
    })
    .optional(),
  groupName: z.string().optional(),
  asAdmin: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAdminRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ADMIN_REMOVE;
  }

  public get description(): string {
    return 'Removes a user or group as site collection administrator';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.userId, o.userName, o.groupId, o.groupName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'userId', 'userName', 'groupId', or 'groupName', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName', 'groupId', 'groupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (!args.options.force) {
        const principalToDelete = args.options.groupId || args.options.groupName ? 'group' : 'user';
        const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove specified ${principalToDelete} from the site administrators list ${args.options.siteUrl}?` });
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
    if (this.verbose) {
      await logger.logToStderr('Removing site administrator as an administrator...');
    }

    const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
    const tenantSiteProperties = await spo.getSiteAdminPropertiesByUrl(args.options.siteUrl, false, logger, this.verbose);
    const primaryAdminLoginName = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, tenantSiteProperties.SiteId, logger, this.verbose);

    if (loginNameToRemove === primaryAdminLoginName) {
      throw 'You cannot remove the primary site collection administrator.';
    }

    const existingAdmins = await this.getSiteAdmins(adminUrl, tenantSiteProperties.SiteId);
    const adminsToSet = existingAdmins.filter(u => u.loginName.toLowerCase() !== loginNameToRemove.toLowerCase());
    await this.setSiteAdminsAsAdmin(adminUrl, tenantSiteProperties.SiteId, adminsToSet);
  }

  private async getSiteAdmins(adminUrl: string, siteId: string): Promise<AdminUserResult[]> {
    const requestOptions: CliRequestOptions = {
      url: `${adminUrl}/_api/SPO.Tenant/GetSiteAdministrators?siteId='${siteId}'`,
      headers: {
        accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;charset=utf-8'
      },
      responseType: 'json'
    };

    const response = await request.post<AdminResult>(requestOptions);
    return response.value;
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
      },
      responseType: 'json'
    };

    await request.post<string>(requestOptions);
  }

  private async callAction(logger: Logger, args: CommandArgs, loginNameToRemove: string): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Removing site administrator...');
    }

    const primaryOwnerLogin = await spo.getPrimaryOwnerLoginFromSite(args.options.siteUrl, logger, this.verbose);
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
}

export default new SpoSiteAdminRemoveCommand();