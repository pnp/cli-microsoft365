import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request from '../../../../request.js';
import { ContextInfo, IdentityResponse, spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { SpoPropertyBagBaseCommand } from '../propertybag/propertybag-base.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoWebReindexCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_REINDEX;
  }

  public get description(): string {
    return 'Requests reindexing the specified subsite';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let requestDigest: string = '';
    let webIdentityResp: IdentityResponse;

    try {
      const res: ContextInfo = await spo.getRequestDigest(args.options.url);
      requestDigest = res.FormDigestValue;

      if (this.debug) {
        await logger.logToStderr(`Retrieved request digest. Retrieving web identity...`);
      }

      const identityResp: IdentityResponse = await spo.getCurrentWebIdentity(args.options.url, requestDigest);
      webIdentityResp = identityResp;

      if (this.debug) {
        await logger.logToStderr(`Retrieved web identity.`);
      }
      if (this.verbose) {
        await logger.logToStderr(`Checking if the site is a no-script site...`);
      }

      const isNoScriptSite: boolean = await SpoPropertyBagBaseCommand.isNoScriptSite(args.options.url, requestDigest, webIdentityResp, logger, this.debug);

      if (isNoScriptSite) {
        if (this.verbose) {
          await logger.logToStderr(`Site is a no-script site. Reindexing lists instead...`);
        }

        await this.reindexLists(args.options.url, requestDigest, logger, webIdentityResp) as any;
      }

      if (this.verbose) {
        await logger.logToStderr(`Site is not a no-script site. Reindexing site...`);
      }

      const requestOptions: any = {
        url: `${args.options.url}/_api/web/allproperties`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      const webProperties: { vti_x005f_searchversion?: number } = await request.get(requestOptions);
      let searchVersion: number = webProperties.vti_x005f_searchversion || 0;
      searchVersion++;

      await SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchVersion.toString(), args.options.url, requestDigest, webIdentityResp, logger, this.debug);
    }
    catch (err: any) {
      this.handleRejectedPromise(err);
    }
  }

  private reindexLists(webUrl: string, requestDigest: string, logger: Logger, webIdentityResp: IdentityResponse): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      (async (): Promise<{ value: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }[] }> => {
        if (this.debug) {
          await logger.logToStderr(`Retrieving information about lists...`);
        }

        const requestOptions: any = {
          url: `${webUrl}/_api/web/lists?$select=NoCrawl,Title,RootFolder/Properties,RootFolder/ServerRelativeUrl&$expand=RootFolder/Properties`,
          headers: {
            'accept': 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        return request.get(requestOptions);
      })()
        .then((lists: { value: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }[] }): Promise<void[]> => {
          const promises: Promise<void>[] = lists.value.map(l => this.reindexList(l, webUrl, requestDigest, webIdentityResp, logger));
          return Promise.all(promises);
        })
        .then((): void => {
          resolve();
        }, (err: any) => reject(err));
    });
  }

  private reindexList(list: { NoCrawl: boolean; Title: string; RootFolder: { Properties: any; ServerRelativeUrl: string; } }, webUrl: string, requestDigest: string, webIdentityResp: IdentityResponse, logger: Logger): Promise<void> {
    return new Promise<void>(async (resolve: () => void, reject: (error: any) => void): Promise<void> => {
      if (list.NoCrawl) {
        if (this.debug) {
          await logger.logToStderr(`List ${list.Title} is excluded from crawling`);
        }
        resolve();
        return;
      }

      spo
        .getFolderIdentity(webIdentityResp.objectIdentity, webUrl, list.RootFolder.ServerRelativeUrl, requestDigest)
        .then((folderIdentityResp: IdentityResponse): Promise<any> => {
          let searchversion: number = list.RootFolder.Properties.vti_x005f_searchversion || 0;
          searchversion++;

          return SpoPropertyBagBaseCommand.setProperty('vti_searchversion', searchversion.toString(), webUrl, requestDigest, folderIdentityResp, logger, this.debug, list.RootFolder.ServerRelativeUrl);
        })
        .then((): void => {
          resolve();
        }, (err: any) => reject(err));
    });
  }
}

export default new SpoWebReindexCommand();