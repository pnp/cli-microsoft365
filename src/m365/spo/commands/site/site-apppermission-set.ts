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
  id: z.string().optional().alias('i'),
  appId: z.string()
    .refine(appId => validation.isValidGuid(appId) === true, {
      error: e => `'${e.input}' is not a valid GUID.`
    })
    .optional(),
  appDisplayName: z.string().optional().alias('n'),
  permission: z.string().pipe(z.enum(['read', 'write', 'manage', 'fullcontrol']).alias('p'))
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAppPermissionSetCommand extends GraphCommand {
  private siteId: string = '';

  public get name(): string {
    return commands.SITE_APPPERMISSION_SET;
  }

  public get description(): string {
    return 'Updates a specific application permission for a site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.id, o.appId, o.appDisplayName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'id', 'appId' or 'appDisplayName', but not multiple.`,
        params: {
          customCode: 'optionSet',
          options: ['id', 'appId', 'appDisplayName']
        }
      });
  }

  private getFilteredPermissions(args: CommandArgs, permissions: Permission[]): Permission[] {
    let filterProperty: string = 'displayName';
    let filterValue: string = args.options.appDisplayName as string;

    if (args.options.appId) {
      filterProperty = 'id';
      filterValue = args.options.appId;
    }

    return permissions.filter((p: Permission) =>
      p.grantedToIdentities!.some(({ application }: IdentitySet) =>
        (application as any)[filterProperty] === filterValue)
    );
  }

  private async getPermission(args: CommandArgs): Promise<string> {
    if (args.options.id) {
      return args.options.id;
    }

    const permissionRequestOptions: any = {
      url: `${this.resource}/v1.0/sites/${this.siteId}/permissions`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response: { value: Permission[] } = await request.get<{ value: Permission[] }>(permissionRequestOptions);

    const sitePermissionItems: Permission[] = this.getFilteredPermissions(args, response.value);

    if (sitePermissionItems.length === 0) {
      throw 'The specified app permission does not exist';
    }

    if (sitePermissionItems.length > 1) {
      throw `Multiple app permissions with displayName ${args.options.appDisplayName} found: ${response.value.map(x => x.grantedToIdentities!.map(y => y.application!.id))}`;
    }

    return sitePermissionItems[0].id!;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      this.siteId = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const sitePermissionId: string = await this.getPermission(args);
      const requestOptions: any = {
        url: `${this.resource}/v1.0/sites/${this.siteId}/permissions/${sitePermissionId}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        data: {
          roles: [args.options.permission]
        },
        responseType: 'json'
      };

      const res = await request.patch(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppPermissionSetCommand();