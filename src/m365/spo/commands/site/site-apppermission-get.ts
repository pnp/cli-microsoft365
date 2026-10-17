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
  id: z.string().alias('i')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAppPermissionGetCommand extends GraphCommand {
  public get name(): string {
    return commands.SITE_APPPERMISSION_GET;
  }

  public get description(): string {
    return 'Gets a specific application permissions for the site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  private getApplicationPermission(args: CommandArgs, siteId: string): Promise<Permission> {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/sites/${siteId}/permissions/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request.get(requestOptions);
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const siteId: string = await spo.getSpoGraphSiteId(args.options.siteUrl);
      const permissionObject: Permission = await this.getApplicationPermission(args, siteId);
      const transposed: { appDisplayName: string; appId: string; permissionId: string, roles: string }[] = [];

      permissionObject.grantedToIdentities!.forEach((permissionEntity: IdentitySet) => {
        transposed.push(
          {
            appDisplayName: permissionEntity.application!.displayName!,
            appId: permissionEntity.application!.id!,
            permissionId: permissionObject.id!,
            roles: permissionObject.roles!.join()
          });
      });

      await logger.log(transposed);

    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAppPermissionGetCommand();
