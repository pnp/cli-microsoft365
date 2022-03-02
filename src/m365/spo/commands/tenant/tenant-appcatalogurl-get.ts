import { Logger } from '../../../../cli';
import request from '../../../../request';
import { spo } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public commandAction(logger: Logger, args: any, cb: (err?: any) => void): void {
    spo
      .getSpoUrl(logger, this.debug)
      .then((spoUrl: string): Promise<string> => {
        const requestOptions: any = {
          url: `${spoUrl}/_api/SP_TenantSettings_Current`,
          headers: {
            accept: 'application/json;odata=nometadata'
          }
        };
    
        return request.get(requestOptions);
      })
      .then((res: string): void => {
        const json = JSON.parse(res);

        if (json.CorporateCatalogUrl) {
          logger.log(json.CorporateCatalogUrl);
        }
        else {
          if (this.verbose) {
            logger.logToStderr("Tenant app catalog is not configured.");
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, logger, cb));
  }
}

module.exports = new SpoTenantAppCatalogUrlGetCommand();