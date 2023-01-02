import { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Page, supportedPageLayouts, supportedPromoteAs } from './Page';
import { Options as spoFileGetOptions } from '../file/file-get';
import { Options as spoListItemSetOptions } from '../listitem/listitem-set';
import * as spoFileGetCommand from '../file/file-get';
import * as spoListItemSetCommand from '../listitem/listitem-set';
import { Cli, CommandOutput } from '../../../../cli/Cli';
import Command from '../../../../Command';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  layoutType?: string;
  promoteAs?: string;
  commentsEnabled: boolean;
  publish: boolean;
  publishMessage?: string;
  description?: string;
  title?: string;
}

class SpoPageAddCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_ADD;
  }

  public get description(): string {
    return 'Creates modern page';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        layoutType: args.options.layoutType,
        promoteAs: args.options.promoteAs,
        commentsEnabled: args.options.commentsEnabled || false,
        publish: args.options.publish || false,
        publishMessage: typeof args.options.publishMessage !== 'undefined',
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-l, --layoutType [layoutType]',
        autocomplete: supportedPageLayouts
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        autocomplete: supportedPromoteAs
      },
      {
        option: '--commentsEnabled'
      },
      {
        option: '--publish'
      },
      {
        option: '--publishMessage [publishMessage]'
      },
      {
        option: '--description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.webUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }

        if (args.options.layoutType &&
          supportedPageLayouts.indexOf(args.options.layoutType) < 0) {
          return `${args.options.layoutType} is not a valid option for layoutType. Allowed values ${supportedPageLayouts.join(', ')}`;
        }

        if (args.options.promoteAs &&
          supportedPromoteAs.indexOf(args.options.promoteAs) < 0) {
          return `${args.options.promoteAs} is not a valid option for promoteAs. Allowed values ${supportedPromoteAs.join(', ')}`;
        }

        if (args.options.promoteAs === 'HomePage' && args.options.layoutType !== 'Home') {
          return 'You can only promote home pages as site home page';
        }

        if (args.options.promoteAs === 'NewsPage' && args.options.layoutType && args.options.layoutType !== 'Article') {
          return 'You can only promote article pages as news article';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const resource = Auth.getResourceFromUrl(args.options.webUrl);
    let requestDigest: string = '';
    let itemId: string = '';
    let pageName: string = args.options.name;
    const serverRelativeSiteUrl: string = urlUtil.getServerRelativeSiteUrl(args.options.webUrl);
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    const pageTitle: string = args.options.title ? args.options.title : (args.options.name.indexOf('.aspx') > -1 ? args.options.name.substr(0, args.options.name.indexOf('.aspx')) : args.options.name);
    let pageId: number | null = null;
    const pageDescription: string = args.options.description || "";

    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const listServerRelativeUrl = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages`;
    const serverRelativeFileUrl: string = `${listServerRelativeUrl}/${pageName}`;

    try {
      const reqDigest = await spo.getRequestDigest(args.options.webUrl);
      requestDigest = reqDigest.FormDigestValue;

      let requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeSiteUrl}/sitepages')/files/AddTemplateFile`,
        headers: {
          'X-RequestDigest': requestDigest,
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          urlOfFile: `${serverRelativeSiteUrl}/sitepages/${pageName}`,
          templateFileType: 3
        },
        responseType: 'json'
      };

      const template = await request.post<{ UniqueId: string }>(requestOptions);
      itemId = template.UniqueId;
      const listItemId = await this.getFileListItemId(args.options.webUrl, serverRelativeFileUrl);
      const layoutType: string = args.options.layoutType || 'Article';

      const listItemSetOptions: spoListItemSetOptions = {
        webUrl: args.options.webUrl,
        listUrl: listServerRelativeUrl,
        id: listItemId,
        systemUpdate: true,
        ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C4118',
        Title: pageTitle,
        ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
        PageLayoutType: layoutType,
        verbose: this.verbose,
        debug: this.debug
      };
      if (args.options.layoutType === 'Article') {
        listItemSetOptions.PromotedState = 0;
        listItemSetOptions.BannerImageUrl = `${resource}/_layouts/15/images/sitepagethumbnail.png, /_layouts/15/images/sitepagethumbnail.png`;
      }
      await Cli.executeCommand(spoListItemSetCommand as Command, { options: { ...listItemSetOptions, _: [] } });

      const pageProps = await Page.checkout(pageName, args.options.webUrl, logger, this.debug, this.verbose);
      if (pageProps) {
        pageId = pageProps.Id;

        bannerImageUrl = pageProps.BannerImageUrl;
        canvasContent1 = pageProps.CanvasContent1;
        layoutWebpartsContent = pageProps.LayoutWebpartsContent;
      }

      if (args.options.promoteAs) {
        const requestOptions: CliRequestOptions = {
          responseType: 'json'
        };

        switch (args.options.promoteAs) {
          case 'HomePage':
            requestOptions.url = `${args.options.webUrl}/_api/web/rootfolder`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.data = {
              WelcomePage: `SitePages/${pageName}`
            };
            await request.post(requestOptions);
            break;
          case 'NewsPage':
            const listItemSetOptions: spoListItemSetOptions = {
              webUrl: args.options.webUrl,
              listUrl: listServerRelativeUrl,
              id: listItemId,
              systemUpdate: true,
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString(),
              verbose: this.verbose,
              debug: this.debug
            };
            await Cli.executeCommand(spoListItemSetCommand as Command, { options: { ...listItemSetOptions, _: [] } });
            break;
          case 'Template':
            requestOptions.url = `${args.options.webUrl}/_api/SitePages/Pages(${listItemId})/SavePageAsTemplate`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'POST',
              'IF-MATCH': '*',
              accept: 'application/json;odata=nometadata'
            };

            const tmpl = await request.post<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string, UniqueId: string }>(requestOptions);

            bannerImageUrl = tmpl.BannerImageUrl;
            canvasContent1 = tmpl.CanvasContent1;
            layoutWebpartsContent = tmpl.LayoutWebpartsContent;
            pageId = tmpl.Id;

            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyid('${tmpl.UniqueId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };

            await request.post(requestOptions);
            break;
        }
      }

      requestOptions = {
        responseType: 'json',
        url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
        headers: {
          'X-RequestDigest': requestDigest,
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        data: {
          BannerImageUrl: bannerImageUrl,
          CanvasContent1: canvasContent1,
          LayoutWebpartsContent: layoutWebpartsContent,
          Description: pageDescription
        }
      };

      await request.post(requestOptions);

      if (args.options.promoteAs === 'Template') {
        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            Title: fileNameWithoutExtension,
            BannerImageUrl: bannerImageUrl,
            CanvasContent1: canvasContent1,
            LayoutWebpartsContent: layoutWebpartsContent,
            Description: pageDescription
          }
        };

        await request.post(requestOptions);
      }

      requestOptions = {
        url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/ListItemAllFields/SetCommentsDisabled(${!args.options.commentsEnabled})`,
        headers: {
          'X-RequestDigest': requestDigest,
          'content-type': 'application/json;odata=nometadata',
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.post(requestOptions);

      if (!args.options.publish) {
        if (args.options.promoteAs !== 'Template' && pageId) {
          requestOptions = {
            responseType: 'json',
            url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
            headers: {
              'content-type': 'application/json;odata=nometadata',
              'accept': 'application/json;odata=nometadata'
            },
            data: {
              Title: pageTitle,
              Description: pageDescription,
              BannerImageUrl: bannerImageUrl,
              CanvasContent1: canvasContent1,
              LayoutWebpartsContent: layoutWebpartsContent
            }
          };
        }
      }
      else {
        requestOptions = {
          url: `${args.options.webUrl}/_api/web/getfilebyid('${itemId}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '').replace(/'/g, '%39')}'&@a2=1`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };
      }

      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
  private async getFileListItemId(webUrl: string, serverRelativeFileUrl: string): Promise<string> {
    const fileGetOptions: spoFileGetOptions = {
      webUrl: webUrl,
      url: serverRelativeFileUrl,
      asListItem: true,
      verbose: this.verbose,
      debug: this.debug
    };
    const fileGetOutput: CommandOutput = await Cli.executeCommandWithOutput(spoFileGetCommand as Command, { options: { ...fileGetOptions, _: [] } });
    const fileGetOutputJson = JSON.parse(fileGetOutput.stdout);
    return fileGetOutputJson.Id;
  }
}

module.exports = new SpoPageAddCommand();
