import {
  ContextInfo, ClientSvcResponse, ClientSvcResponseContents
} from '../../spo';
import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandError
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { OrgAssetsResponse, OrgAssets } from './OrgAssets';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
}

class SpoOrgNewsSiteListCommand extends SpoCommand {
  public get name(): string {
    return `${commands.ORGASSETSLIBRARY_LIST}`;
  }

  public get description(): string {
    return 'List all libraries that are assigned as asset library';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Method Name="GetOrgAssets" Id="6" ObjectPathId="3" /></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];

        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          const orgAssetsResponse: OrgAssetsResponse = json[json.length - 1];

          if (orgAssetsResponse === null || orgAssetsResponse.OrgAssetsLibraries === undefined) {
            cmd.log("No libraries in Organization Assets");
          } else {
            const orgAssets: OrgAssets = {
              Url: orgAssetsResponse.Url.DecodedUrl,
              Libraries: orgAssetsResponse.OrgAssetsLibraries._Child_Items_.map(t => {
                return {
                  DisplayName: t.DisplayName,
                  LibraryUrl: t.LibraryUrl.DecodedUrl,
                  ListId: t.ListId,
                  ThumbnailUrl: t.ThumbnailUrl != null ? t.ThumbnailUrl.DecodedUrl : null
                }
              })
            }

            if (args.options.output === 'json') {
              cmd.log(JSON.stringify(orgAssets));
            } else {
              cmd.log(orgAssets);
            }

            if (this.verbose) {
              cmd.log(chalk.green('DONE'));
            }
          }
          cb();
        }
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const parentOptions: CommandOption[] = super.options();
    return parentOptions;
  }
}

module.exports = new SpoOrgNewsSiteListCommand();
