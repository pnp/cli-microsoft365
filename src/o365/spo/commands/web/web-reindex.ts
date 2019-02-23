import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import { SpoPropertyBagBaseCommand } from '../propertybag/propertybag-base';
import { ContextInfo } from '../../spo';
import { ClientSvc, IdentityResponse } from '../../common/ClientSvc';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebReindexCommand extends SpoCommand {
  private reindexedLists: boolean;

  constructor() {
    super()/* istanbul ignore next */;
    this.reindexedLists = false;
  }

  public get name(): string {
    return commands.WEB_REINDEX;
  }

  public get description(): string {
    return 'Requests reindexing the specified subsite';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    let webIdentityResp: IdentityResponse;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
        }

        siteAccessToken = accessToken;

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): Promise<IdentityResponse> => {
        requestDigest = res.FormDigestValue;

        if (this.debug) {
          cmd.log(`Retrieved request digest. Retrieving web identity...`);
        }

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, siteAccessToken, requestDigest);
      })
      .then((identityResp: IdentityResponse): Promise<boolean> => {
        webIdentityResp = identityResp;

        if (this.debug) {
          cmd.log(`Retrieved web identity.`);
        }
        if (this.verbose) {
          cmd.log(`Checking if the site is a no-script site...`);
        }

        return SpoPropertyBagBaseCommand.isNoScriptSite(args.options.webUrl, requestDigest, siteAccessToken, webIdentityResp, clientSvcCommons);
      })
      .then((isNoScriptSite: boolean): request.RequestPromise | Promise<void> => {
        if (isNoScriptSite) {
          if (this.verbose) {
            cmd.log(`Site is a no-script site. Reindexing lists instead...`);
          }

          return this.reindexLists(args.options.webUrl, requestDigest, siteAccessToken, cmd, webIdentityResp, clientSvcCommons);
        }

        if (this.verbose) {
          cmd.log(`Site is not a no-script site. Reindexing site...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/allproperties`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((webProperties: { vti_x005f_searchversion?: number }): Promise<any> => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(webProperties);
          cmd.log('');
        }

        let searchVersion: number = webProperties.vti_x005f_searchversion || 0;
        searchVersion++;

        return SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchVersion.toString(), args.options.webUrl, requestDigest, siteAccessToken, webIdentityResp, cmd, this.debug);
      })
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');
        }

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => {
        if (this.reindexedLists) {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }
        else {
          this.handleRejectedPromise(err, cmd, cb);
        }
      });
  }

  private reindexLists(webUrl: string, requestDigest: string, siteAccessToken: string, cmd: CommandInstance, webIdentityResp: IdentityResponse, clientSvcCommons: ClientSvc): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      ((): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieving information about lists...`);
        }

        const requestOptions: any = {
          url: `${webUrl}/_api/web/lists?$select=NoCrawl,Title,RootFolder/Properties,RootFolder/ServerRelativeUrl&$expand=RootFolder/Properties`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })()
        .then((lists: { value: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }[] }): Promise<void[]> => {
          const promises: Promise<void>[] = lists.value.map(l => this.reindexList(l, webUrl, requestDigest, siteAccessToken, webIdentityResp, clientSvcCommons, cmd));
          return Promise.all(promises);
        })
        .then((): void => {
          this.reindexedLists = true;
          reject(undefined);
        }, (err: any) => reject(err));
    });
  }

  private reindexList(list: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }, webUrl: string, requestDigest: string, accessToken: string, webIdentityResp: IdentityResponse, clientSvcCommons: ClientSvc, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (list.NoCrawl) {
        if (this.debug) {
          cmd.log(`List ${list.Title} is excluded from crawling`);
        }
        resolve();
        return;
      }

      clientSvcCommons
        .getFolderIdentity(webIdentityResp.objectIdentity, webUrl, list.RootFolder.ServerRelativeUrl, accessToken, requestDigest)
        .then((folderIdentityResp: IdentityResponse): Promise<any> => {
          let searchversion: number = list.RootFolder.Properties.vti_x005f_searchversion || 0;
          searchversion++;

          return SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchversion.toString(), webUrl, requestDigest, accessToken, folderIdentityResp, cmd, this.debug, list.RootFolder.ServerRelativeUrl);
        })
        .then((): void => {
          resolve();
        }, (err: any) => reject(err));
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the subsite to reindex'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To request reindexing a subsite, you have to first log in to SharePoint
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the subsite to be reindexed is a no-script site, the command will request
    reindexing all lists from the subsite that haven't been excluded from the
    search index.
        
  Examples:
  
    Request reindexing the subsite ${chalk.grey('https://contoso.sharepoint.com/subsite')}
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/subsite
      `);
  }
}

module.exports = new SpoWebReindexCommand();