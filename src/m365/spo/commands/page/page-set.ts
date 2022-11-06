import { Auth } from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { spo } from '../../../../utils/spo';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { Page, supportedPageLayouts, supportedPromoteAs } from './Page';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  layoutType?: string;
  promoteAs?: string;
  commentsEnabled?: string;
  publish: boolean;
  publishMessage?: string;
  description?: string;
  title?: string;
}

class SpoPageSetCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_SET;
  }

  public get description(): string {
    return 'Updates modern page properties';
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
        description: typeof args.options.description !== 'undefined',
        title: typeof args.options.title !== 'undefined'
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
        option: '-l, --layoutType [layoutType]',
        autocomplete: supportedPageLayouts
      },
      {
        option: '-p, --promoteAs [promoteAs]',
        autocomplete: supportedPromoteAs
      },
      {
        option: '--commentsEnabled [commentsEnabled]',
        autocomplete: ['true', 'false']
      },
      {
        option: '--publish'
      },
      {
        option: '--publishMessage [publishMessage]'
      },
      {
        option: '--description [description]'
      },
      {
        option: '--title [title]'
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

        if (typeof args.options.commentsEnabled !== 'undefined' &&
          args.options.commentsEnabled !== 'true' &&
          args.options.commentsEnabled !== 'false') {
          return `${args.options.commentsEnabled} is not a valid value for commentsEnabled. Allowed values true|false`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let pageName: string = args.options.name;
    const fileNameWithoutExtension: string = pageName.replace('.aspx', '');
    let bannerImageUrl: string = '';
    let canvasContent1: string = '';
    let layoutWebpartsContent: string = '';
    let pageTitle: string = args.options.title || "";
    let pageId: number | null = null;
    let pageDescription: string = args.options.description || "";
    let topicHeader: string = "";
    let authorByline: string[] = [];
    const pageData: any = {};

    if (!pageName.endsWith('.aspx')) {
      pageName += '.aspx';
    }
    const serverRelativeFileUrl: string = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages/${pageName}`;
    const needsToSavePage = !!args.options.title || !!args.options.description;

    try {
      const requestDigestResult = await spo.getRequestDigest(args.options.webUrl);
      const requestDigest = requestDigestResult.FormDigestValue;
      const page = await Page.checkout(args.options.name, args.options.webUrl, logger, this.debug, this.verbose);

      if (page) {
        pageTitle = pageTitle || page.Title;
        pageId = page.Id;

        bannerImageUrl = page.BannerImageUrl;
        canvasContent1 = page.CanvasContent1;
        layoutWebpartsContent = page.LayoutWebpartsContent;
        pageDescription = pageDescription || page.Description;
        topicHeader = page.TopicHeader;
        authorByline = page.AuthorByline;
      }

      if (args.options.layoutType) {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            accept: 'application/json;odata=nometadata'
          },
          data: {
            PageLayoutType: args.options.layoutType
          },
          responseType: 'json'
        };

        if (args.options.layoutType === 'Article') {
          requestOptions.data.PromotedState = 0;
          requestOptions.data.BannerImageUrl = {
            Description: '/_layouts/15/images/sitepagethumbnail.png',
            Url: `${resource}/_layouts/15/images/sitepagethumbnail.png`
          };
        }

        await request.post(requestOptions);
      }

      if (args.options.promoteAs) {
        const requestOptions: any = {
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
            break;
          case 'NewsPage':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.data = {
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString().replace('Z', '')
            };
            break;
          case 'Template':
            requestOptions.url = `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields`;
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              accept: 'application/json;odata=nometadata'
            };
            break;
        }

        const pageRes = await request.post<{ Id: string }>(requestOptions);

        if (args.options.promoteAs === 'Template') {
          const requestOptions: any = {
            responseType: 'json',
            url: `${args.options.webUrl}/_api/SitePages/Pages(${pageRes.Id})/SavePageAsTemplate`,
            headers: {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'POST',
              'IF-MATCH': '*',
              accept: 'application/json;odata=nometadata'
            }
          };

          const res = await request.post<{ Id: number | null, BannerImageUrl: string, CanvasContent1: string, LayoutWebpartsContent: string }>(requestOptions);
          if (fileNameWithoutExtension) {
            pageData.Title = fileNameWithoutExtension;
          }
          if (pageDescription) {
            pageData.Description = pageDescription;
          }
          if (res.BannerImageUrl) {
            pageData.BannerImageUrl = res.BannerImageUrl;
          }
          if (res.LayoutWebpartsContent) {
            pageData.LayoutWebpartsContent = res.LayoutWebpartsContent;
          }
          if (res.CanvasContent1) {
            pageData.CanvasContent1 = res.CanvasContent1;
          }

          pageId = res.Id;
        }
      }
      if (args.options.promoteAs !== 'Template') {
        if (pageTitle) {
          pageData.Title = pageTitle;
        }
        if (pageDescription) {
          pageData.Description = pageDescription;
        }
        if (bannerImageUrl) {
          pageData.BannerImageUrl = bannerImageUrl;
        }
        if (canvasContent1) {
          pageData.CanvasContent1 = canvasContent1;
        }
        if (layoutWebpartsContent) {
          pageData.LayoutWebpartsContent = layoutWebpartsContent;
        }
        if (topicHeader) {
          pageData.TopicHeader = topicHeader;
        }
        if (authorByline) {
          pageData.AuthorByline = authorByline;
        }
      }

      if (needsToSavePage) {
        const requestOptions: any = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePage`,
          headers: {
            'X-RequestDigest': requestDigest,
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          data: pageData
        };

        await request.post(requestOptions);
      }

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
          data: pageData
        };

        await request.post(requestOptions);
      }

      if (typeof args.options.commentsEnabled !== 'undefined') {
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === 'false'})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }

      let requestOptions: any;

      if (!args.options.publish) {
        if (args.options.promoteAs === 'Template' || !pageId) {
          return;
        }

        requestOptions = {
          responseType: 'json',
          url: `${args.options.webUrl}/_api/SitePages/Pages(${pageId})/SavePageAsDraft`,
          headers: {
            'content-type': 'application/json;odata=nometadata',
            'accept': 'application/json;odata=nometadata'
          },
          data: pageData
        };
      }
      else {
        requestOptions = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '').replace(/'/g, '%39')}'&@a2=1`,
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
}

module.exports = new SpoPageSetCommand();
