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

class SpoTenantDeletedsiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_DELETEDSITE_LIST;
  }

  public get description(): string {
    return 'Lists the deleted site collections in the tenant Recycle Bin.';
  }

  protected requiresTenantAdmin(): boolean {
    return true;
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
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="32" ObjectPathId="31" /><ObjectPath Id="34" ObjectPathId="33" /><Query Id="35" ObjectPathId="33"><Query SelectAllProperties="true"><Properties><Property Name="NextStartIndexFromSharePoint" ScalarProperty="true" /></Properties></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="31" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="33" ParentId="31" Name="GetDeletedSitePropertiesFromSharePoint"><Parameters><Parameter Type="Null" /></Parameters></Method></ObjectPaths></Request>`
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

        let result;
        if (json[6] && json[6]['_Child_Items_']) {
          result = json[6]['_Child_Items_'] as any[];
        }

        if (result && result.length > 0) {
          result.forEach(c => {
            delete c['_ObjectIdentity_'];
            delete c['_ObjectType_'];
            c.SiteId = c.SiteId.replace('/Guid(', '').replace(')/', '');
            const dateChunks: number[] = (c.DeletionTime as string)
              .replace('/Date(', '')
              .replace(')/', '')
              .split(',')
              .map(c => {
                return parseInt(c);
              });
            c.DeletionTime = new Date(dateChunks[0], dateChunks[1], dateChunks[2], dateChunks[3], dateChunks[4], dateChunks[5], dateChunks[6]).toISOString();
          });

          if (args.options.output === 'json') {
            cmd.log(result);
          }
          else {
            cmd.log(result.map(e => {
              return {
                Url: e.Url,
                StorageMaximumLevel: e.StorageMaximumLevel,
                UserCodeMaximumLevel: e.UserCodeMaximumLevel,
                DeletionTime: e.DeletionTime,
                DaysRemaining: e.DaysRemaining
              };
            }));
            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No deleted site collections found');
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
  
  Lists the deleted site collections
      ${chalk.grey(config.delimiter)} ${commands.TENANT_DELETEDSITE_LIST}
  ` );
  }
}

module.exports = new SpoTenantDeletedsiteListCommand();