import { Permission } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface AppInfo {
  appId: string;
  displayName: string;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  permission: z.string().pipe(z.enum(['read', 'write', 'manage', 'fullcontrol']).alias('p')),
  appId: z.string()
    .refine(appId => validation.isValidGuid(appId) === true, {
      error: e => `'${e.input}' is not a valid GUID.`
    })
    .optional()
    .alias('i'),
  appDisplayName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAppPermissionAddCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_ADD;
  }

  public get description(): string {
    return 'Adds an application permissions to the site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => o.appId !== undefined || o.appDisplayName !== undefined, {
        error: `Specify either 'appId' or 'appDisplayName'.`,
        params: {
          customCode: 'optionSet',
          options: ['appId', 'appDisplayName']
        }
      });
  }

  private async getAppInfo(args: CommandArgs): Promise<AppInfo> {
    if (args.options.appId && args.options.appDisplayName) {
      return {
        appId: args.options.appId as string,
        displayName: args.options.appDisplayName as string
      };
    }

    let endpoint: string;

    if (args.options.appId) {
      endpoint = `${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,displayName&$filter=appId eq '${formatting.encodeQueryParameter(args.options.appId as string)}'`;
    }
    else {
      endpoint = `${this.resource}/v1.0/myorganization/servicePrincipals?$select=appId,displayName&$filter=displayName eq '${formatting.encodeQueryParameter(args.options.appDisplayName as string)}'`;
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
      throw "The specified Microsoft Entra app does not exist";
    }

    if (response.value.length > 1) {
      throw `Multiple Microsoft Entra apps with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.appId)}`;
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

      await logger.log(permission);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppPermissionAddCommand();