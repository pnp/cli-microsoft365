import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import { SpoPropertyBagBaseCommand } from '../propertybag/propertybag-base';
import { ContextInfo } from '../../spo';
import { ClientSvc, IdentityResponse } from '../../ClientSvc';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoWebReindexCommand extends SpoCommand {
  private reindexedLists: boolean;

  constructor() {
    super();
    this.reindexedLists = false;
  }

  public get name(): string {
    return commands.WEB_REINDEX;
  }

  public get description(): string {
    return 'Requests reindexing the specified subsite';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const clientSvcCommons: ClientSvc = new ClientSvc(cmd, this.debug);
    let requestDigest: string = '';
    let webIdentityResp: IdentityResponse;

    this
      .getRequestDigest(args.options.webUrl)
      .then((res: ContextInfo): Promise<IdentityResponse> => {
        requestDigest = res.FormDigestValue;

        if (this.debug) {
          cmd.log(`Retrieved request digest. Retrieving web identity...`);
        }

        return clientSvcCommons.getCurrentWebIdentity(args.options.webUrl, requestDigest);
      })
      .then((identityResp: IdentityResponse): Promise<boolean> => {
        webIdentityResp = identityResp;

        if (this.debug) {
          cmd.log(`Retrieved web identity.`);
        }
        if (this.verbose) {
          cmd.log(`Checking if the site is a no-script site...`);
        }

        return SpoPropertyBagBaseCommand.isNoScriptSite(args.options.webUrl, requestDigest, webIdentityResp, clientSvcCommons);
      })
      .then((isNoScriptSite: boolean): Promise<{ vti_x005f_searchversion?: number }> => {
        if (isNoScriptSite) {
          if (this.verbose) {
            cmd.log(`Site is a no-script site. Reindexing lists instead...`);
          }

          return this.reindexLists(args.options.webUrl, requestDigest, cmd, webIdentityResp, clientSvcCommons) as any;
        }

        if (this.verbose) {
          cmd.log(`Site is not a no-script site. Reindexing site...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/allproperties`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })
      .then((webProperties: { vti_x005f_searchversion?: number }): Promise<any> => {
        let searchVersion: number = webProperties.vti_x005f_searchversion || 0;
        searchVersion++;

        return SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchVersion.toString(), args.options.webUrl, requestDigest, webIdentityResp, cmd, this.debug);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => {
        if (this.reindexedLists) {
          if (this.verbose) {
            cmd.log(chalk.green('DONE'));
          }

          cb();
        }
        else {
          this.handleRejectedPromise(err, cmd, cb);
        }
      });
  }

  private reindexLists(webUrl: string, requestDigest: string, cmd: CommandInstance, webIdentityResp: IdentityResponse, clientSvcCommons: ClientSvc): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      ((): Promise<{ value: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }[] }> => {
        if (this.debug) {
          cmd.log(`Retrieving information about lists...`);
        }

        const requestOptions: any = {
          url: `${webUrl}/_api/web/lists?$select=NoCrawl,Title,RootFolder/Properties,RootFolder/ServerRelativeUrl&$expand=RootFolder/Properties`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.get(requestOptions);
      })()
        .then((lists: { value: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }[] }): Promise<void[]> => {
          const promises: Promise<void>[] = lists.value.map(l => this.reindexList(l, webUrl, requestDigest, webIdentityResp, clientSvcCommons, cmd));
          return Promise.all(promises);
        })
        .then((): void => {
          this.reindexedLists = true;
          reject(undefined);
        }, (err: any) => reject(err));
    });
  }

  private reindexList(list: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }, webUrl: string, requestDigest: string, webIdentityResp: IdentityResponse, clientSvcCommons: ClientSvc, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      if (list.NoCrawl) {
        if (this.debug) {
          cmd.log(`List ${list.Title} is excluded from crawling`);
        }
        resolve();
        return;
      }

      clientSvcCommons
        .getFolderIdentity(webIdentityResp.objectIdentity, webUrl, list.RootFolder.ServerRelativeUrl, requestDigest)
        .then((folderIdentityResp: IdentityResponse): Promise<any> => {
          let searchversion: number = list.RootFolder.Properties.vti_x005f_searchversion || 0;
          searchversion++;

          return SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchversion.toString(), webUrl, requestDigest, folderIdentityResp, cmd, this.debug, list.RootFolder.ServerRelativeUrl);
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
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoWebReindexCommand();