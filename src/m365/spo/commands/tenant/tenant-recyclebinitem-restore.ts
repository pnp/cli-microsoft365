import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
  wait: boolean;
}

class SpoTenantRecycleBinItemRestoreCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_RECYCLEBINITEM_RESTORE;
  }

  public get description(): string {
    return 'Restores the specified deleted site collection from tenant recycle bin';
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
        wait: args.options.wait
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.siteUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('siteUrl');
    this.types.boolean.push('wait');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.wait) {
      await this.warn(logger, `Option 'wait' is deprecated and will be removed in the next major release.`);
    }

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

      // Here, we return a fixed response due to removing the '--wait' functionality as it is deprecated.
      // This has to be removed in the next major release.
      await logger.log({
        HasTimedout: false,
        IsComplete: !!args.options.wait,
        PollingInterval: 15000
      });
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