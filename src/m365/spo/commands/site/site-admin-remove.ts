import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { AdminResult, AdminUserResult } from './SiteAdmin.js';

interface CommandArgs {
  options: Options;
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
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        force: !!args.options.force,
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

  #initTypes(): void {
    this.types.string.push('siteUrl', 'userId', 'userName', 'groupId', 'groupName');
    this.types.boolean.push('force', 'asAdmin');
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
    const siteId = await this.getSiteId(args.options.siteUrl, logger);
    const primaryAdminLoginName = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, siteId, logger, this.verbose);

    if (loginNameToRemove === primaryAdminLoginName) {
      throw 'You cannot remove the primary site collection administrator.';
    }

    const existingAdmins = await this.getSiteAdmins(adminUrl, siteId);
    const adminsToSet = existingAdmins.filter(u => u.loginName.toLowerCase() !== loginNameToRemove.toLowerCase());
    await this.setSiteAdminsAsAdmin(adminUrl, siteId, adminsToSet);
  }

  private async getSiteId(siteUrl: string, logger: Logger): Promise<string> {
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