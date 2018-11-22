import auth from '../../SpoAuth';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import SpoCommand from '../../SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  scope?: string;
  siteUrl?: string;
}

export abstract class SpoAppBaseCommand extends SpoCommand {

  public getAppCatalogSiteUrl(cmd: CommandInstance, args: CommandArgs): Promise<string> {

    if (args.options.scope === 'sitecollection') {
      
      return Promise.resolve(args.options.siteUrl as string);
    }

    if (args.options.appCatalogUrl) {

      return Promise.resolve(args.options.appCatalogUrl);
    }

    return this.getTenantAppCatalogUrl(cmd);
  }

  protected getTenantAppCatalogUrl(cmd: CommandInstance): Promise<string> {
    return new Promise<string>((resolve: (appCatalogUrl: string) => void, reject: (error: any) => void): void => {

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
            reject(response.ErrorInfo.ErrorMessage);
            return;
          }
          else {
            const catalogUrl: string = json.pop().CorporateCatalogUrl;

            if (catalogUrl) {
              resolve(catalogUrl);
            }
            else {
              reject("Tenant app catalog is not configured.");
            }
          }
        }, (err: any): void => reject(err));
    });
  }
}