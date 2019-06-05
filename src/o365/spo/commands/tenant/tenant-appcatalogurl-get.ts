import request from '../../../../request';
import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    this
      .getSpoUrl(cmd, this.debug)
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
          cmd.log(json.CorporateCatalogUrl);
        }
        else {
          if (this.verbose) {
            cmd.log("Tenant app catalog is not configured.");
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Get the URL of the tenant app catalog
      ${commands.TENANT_APPCATALOGURL_GET}
  ` );
  }
}

module.exports = new SpoTenantAppCatalogUrlGetCommand();