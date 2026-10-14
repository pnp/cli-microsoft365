import { IdentitySet, Permission } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  appId: z.string().optional().alias('i'),
  appDisplayName: z.string().optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAppPermissionListCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_LIST;
  }

  public get description(): string {
    return 'Lists application permissions for a site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => !(o.appId !== undefined && o.appDisplayName !== undefined), {
        error: `Provide either appId or appDisplayName, not both.`
      });
  }

  private getFilteredPermissions(args: CommandArgs, permissions: Permission[]): Permission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: Permission) => {
      return p.grantedToIdentities!.some(({ application }: IdentitySet) =>
        (application as any)[filterProperty] === filterValue);
    });
  }

  private getApplicationPermission(permissionId: string): Promise<Permission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${permissionId}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get<Permission>(requestOptions);
  }

  private getTransposed(permissions: Permission[]): { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] {
    const transposed: { appDisplayName: string; appId: string; permissionId: string, roles: string[] }[] = [];

    permissions.forEach((permissionObject: Permission) => {
      permissionObject.grantedToIdentities!.forEach((permissionEntity: IdentitySet) => {
        transposed.push(
          {
            appDisplayName: permissionEntity.application!.displayName!,
            appId: permissionEntity.application!.id!,
            permissionId: permissionObject.id!,
            roles: permissionObject.roles!
          });
      });
    });

    return transposed;
  }

  private getPermissions(): Promise<{ value: Permission[] }> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const permRes: { value: Permission[] } = await this.getPermissions();
      let permissions: Permission[] = permRes.value;

      if (args.options.appId || args.options.appDisplayName) {
        permissions = this.getFilteredPermissions(args, permRes.value);
      }

      const res: Permission[] = await Promise.all(permissions.map(g => this.getApplicationPermission(g.id!)));
      await logger.log(this.getTransposed(res));

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppPermissionListCommand();