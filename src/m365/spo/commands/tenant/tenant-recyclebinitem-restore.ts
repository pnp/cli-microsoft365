import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string().refine(val => validation.isValidSharePointUrl(val) === true, { message: 'The value is not a valid SharePoint site URL.' }).alias('u')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores the specified deleted site collection from tenant recycle bin';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Restoring site collection '${args.options.siteUrl}' from recycle bin.`);
      }

      const siteUrl = urlUtil.removeTrailingSlashes(args.options.siteUrl);
      const adminUrl: string = await spo.getSpoAdminUrl(logger, this.debug);
      const requestOptions: CliRequestOptions = {
        url: `${adminUrl}/_api/SPO.Tenant/RestoreDeletedSite`,
        headers: {
          accept: 'application/json;odata=nometadata',
          'content-type': 'application/json;charset=utf-8'
        },
        data: { siteUrl },
        responseType: 'json'
      };

      await request.post(requestOptions);

      const groupId = await this.getSiteGroupId(adminUrl, siteUrl);

      if (groupId && groupId !== '00000000-0000-0000-0000-000000000000') {
        if (this.verbose) {
          await logger.logToStderr(`Restoring Microsoft 365 group with ID '${groupId}' from recycle bin.`);
        }

        const restoreOptions: CliRequestOptions = {
          url: `https://graph.microsoft.com/v1.0/directory/deletedItems/${groupId}/restore`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'content-type': 'application/json'
          },
          responseType: 'json'
        };

        await request.post(restoreOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getSiteGroupId(adminUrl: string, url: string): Promise<string | undefined> {
    const sites = await odata.getAllItems<{ GroupId?: string }>(`${adminUrl}/_api/web/lists/GetByTitle('DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS')/items?$filter=SiteUrl eq '${formatting.encodeQueryParameter(url)}'&$select=GroupId`);
    return sites[0].GroupId;
  }
}

export default new SpoTenantRecycleBinItemRestoreCommand();