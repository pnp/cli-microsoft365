import { Permission } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface AppInfo {
  appId: string;
  displayName: string;
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  permission: string;
  appId?: string;
  appDisplayName?: string;
}

class SpoSiteAppPermissionAddCommand extends GraphCommand {
  private siteId: string = '';
  private roles: string[] = ['read', 'write', 'manage', 'fullcontrol'];

  public get name(): string {
    return commands.SITE_APPPERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds an application permissions to the site';
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
        appId: typeof args.options.appId !== 'undefined',
        appDisplayName: typeof args.options.appDisplayName !== 'undefined',
        permissions: args.options.permissions
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '-p, --permission <permission>',
        autocomplete: this.roles
      },
      {
        option: '-i, --appId [appId]'
      },
      {
        option: '-n, --appDisplayName [appDisplayName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.appId && !validation.isValidGuid(args.options.appId)) {
          return `${args.options.appId} is not a valid GUID`;
        }

        if (this.roles.indexOf(args.options.permission) === -1) {
          return `${args.options.permission} is not a valid permission value. Allowed values are ${this.roles.join('|')}`;
        }

        return validation.isValidSharePointUrl(args.options.siteUrl);
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['appId', 'appDisplayName'] });
  }

  private async getAppInfo(args: CommandArgs): Promise<AppInfo> {
    if (args.options.appId && args.options.appDisplayName) {
      return {
        appId: args.options.appId as string,
        displayName: args.options.appDisplayName as string
      };
    }

    let endpoint: string = "";

    if (args.options.appId) {
      endpoint = `${this.resource}/v1.0/myorganization/applications?$filter=appId eq '${formatting.encodeQueryParameter(args.options.appId as string)}'`;
    }
    else {
      endpoint = `${this.resource}/v1.0/myorganization/applications?$filter=displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
    }

    const appRequestOptions: any = {
      url: endpoint,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: { value: AppInfo[] } = await request.get<{ value: AppInfo[] }>(appRequestOptions);

    const appItem: AppInfo | undefined = response.value[0];

    if (!appItem) {
      throw "The specified Azure AD app does not exist";
    }

    if (response.value.length > 1) {
      throw `Multiple Azure AD app with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.appId)}`;
    }

    return {
      appId: appItem.appId,
      displayName: appItem.displayName
    };
  }

  /**
   * Checks if the requested permission needs elevation after the initial creation.
   */
  private roleNeedsElevation(permission: string): boolean {
    return ['manage', 'fullcontrol'].indexOf(permission) > -1;
  }

  /**
   * Grants the app 'read' or 'write' permissions to the site.
   * 
   * Explanation:
   * 'manage' and 'fullcontrol' permissions cannot be granted directly when adding app permissions.
   * They can currently only be assigned when updating existing app permissions.
   * We therefore assign 'write' permissions first, and update it to the requested role afterwards.  
   */
  private addPermissions(args: CommandArgs, appInfo: AppInfo): Promise<Permission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      data: {
        roles: [this.roleNeedsElevation(args.options.permission) ? 'write' : args.options.permission],
        grantedToIdentities: [{ application: { "id": appInfo.appId, "displayName": appInfo.displayName } }]
      },
      responseType: 'json'
    };

    return request.post(requestOptions);
  }

  /**
   * Updates the granted permissions to 'manage' or 'fullcontrol'.
   */
  private elevatePermissions(args: CommandArgs, permission: Permission): Promise<Permission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permission.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      data: {
        roles: [args.options.permission]
      },
      responseType: 'json'
    };

    return request.patch(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const appInfo: AppInfo = await this.getAppInfo(args);
      let permission = await this.addPermissions(args, appInfo);

      if (this.roleNeedsElevation(args.options.permission)) {
        permission = await this.elevatePermissions(args, permission);
      }

      logger.log(permission);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppPermissionAddCommand();