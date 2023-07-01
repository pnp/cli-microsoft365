import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      const requestOptions: any = {
        url: `${spoUrl}/_api/SP_TenantSettings_Current`,
        headers: {
          accept: 'application/json;odata=nometadata'
        }
      };

      const res: string = await request.get(requestOptions);
      const json = JSON.parse(res);

      if (json.CorporateCatalogUrl) {
        await logger.log(json.CorporateCatalogUrl);
      }
      else {
        if (this.verbose) {
          await logger.logToStderr("Tenant app catalog is not configured.");
        }
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoTenantAppCatalogUrlGetCommand();