import auth from '../../SpoAuth';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import * as request from 'request-promise-native';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import {
  CommandError
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        return this.getRequestDigest(cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_vti_bin/client.svc/ProcessQuery`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'X-RequestDigest': res.FormDigestValue
          }),
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="58" ObjectPathId="57" /><Query Id="59" ObjectPathId="57"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticProperty Id="57" TypeId="{e9a11c41-0667-4c14-a4a5-e0d6cf67f6fa}" Name="Current" /></ObjectPaths></Request>`
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          const catalogUrl: string = json.pop().CorporateCatalogUrl;

          if (catalogUrl) {
            cmd.log(catalogUrl);
          }
          else {
            if (this.verbose) {
              cmd.log("Tenant app catalog is not configured.");
            }
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online
    tenant admin site, using the ${chalk.blue(commands.LOGIN)} command.

  Examples:
  
    Get the URL of the tenant app catalog
      ${chalk.grey(config.delimiter)} ${commands.TENANT_APPCATALOGURL_GET}
  ` );
  }
}

module.exports = new SpoTenantAppCatalogUrlGetCommand();