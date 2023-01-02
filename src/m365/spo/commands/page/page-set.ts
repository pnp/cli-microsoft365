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
  commentsEnabled?: boolean;
  publish: boolean;
  publishMessage?: string;
  description?: string;
  title?: string;
  demoteFrom?: string;
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
    this.#initTypes();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        layoutType: args.options.layoutType || false,
        promoteAs: args.options.promoteAs || false,
        demotefrom: args.options.demoteFrom || false,
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
        option: '--demoteFrom [demoteFrom]',
        autocomplete: ['NewsPage']
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

  #initTypes(): void {
    this.types.boolean.push('commentsEnabled');
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

        if (args.options.demoteFrom &&
          args.options.demoteFrom !== 'NewsPage') {
          return `${args.options.demoteFrom} is not a valid option for demoteFrom. The only allowed value is 'NewsPage'`;
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
    const listServerRelativeUrl = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/sitepages`;
    const serverRelativeFileUrl: string = `${listServerRelativeUrl}/${pageName}`;
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
        const itemId = await this.getFileListItemId(args.options.webUrl, serverRelativeFileUrl);
        const listItemSetOptions: spoListItemSetOptions = {
          webUrl: args.options.webUrl,
          listUrl: listServerRelativeUrl,
          id: itemId,
          systemUpdate: true,
          PageLayoutType: args.options.layoutType,
          verbose: this.verbose,
          debug: this.debug
        };
        if (args.options.layoutType === 'Article') {
          listItemSetOptions.PromotedState = 0;
          listItemSetOptions.BannerImageUrl = `${resource}/_layouts/15/images/sitepagethumbnail.png, /_layouts/15/images/sitepagethumbnail.png`;
        }
        await Cli.executeCommand(spoListItemSetCommand as Command, { options: { ...listItemSetOptions, _: [] } });
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
            const newsPageItemId = await this.getFileListItemId(args.options.webUrl, serverRelativeFileUrl);
            const listItemSetOptions: spoListItemSetOptions = {
              webUrl: args.options.webUrl,
              listUrl: listServerRelativeUrl,
              id: newsPageItemId,
              systemUpdate: true,
              PromotedState: 2,
              FirstPublishedDate: new Date().toISOString(),
              verbose: this.verbose,
              debug: this.debug
            };
            await Cli.executeCommand(spoListItemSetCommand as Command, { options: { ...listItemSetOptions, _: [] } });
            break;
          case 'Template':
            const templateItemId = await this.getFileListItemId(args.options.webUrl, serverRelativeFileUrl);
            requestOptions.headers = {
              'X-RequestDigest': requestDigest,
              'content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'POST',
              'IF-MATCH': '*',
              accept: 'application/json;odata=nometadata'
            };
            requestOptions.url = `${args.options.webUrl}/_api/SitePages/Pages(${templateItemId})/SavePageAsTemplate`;
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
            break;
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
        const requestOptions: CliRequestOptions = {
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
        const requestOptions: CliRequestOptions = {
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
        const requestOptions: CliRequestOptions = {
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/ListItemAllFields/SetCommentsDisabled(${args.options.commentsEnabled === false})`,
          headers: {
            'X-RequestDigest': requestDigest,
            'content-type': 'application/json;odata=nometadata',
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }

      if (args.options.demoteFrom === 'NewsPage') {
        const fileId = await this.getFileListItemId(args.options.webUrl, serverRelativeFileUrl);
        const listItemSetOptions: spoListItemSetOptions = {
          webUrl: args.options.webUrl,
          listUrl: listServerRelativeUrl,
          id: fileId,
          systemUpdate: true,
          PromotedState: 0,
          verbose: this.verbose,
          debug: this.debug
        };
        await Cli.executeCommandWithOutput(spoListItemSetCommand as Command, { options: { ...listItemSetOptions, _: [] } });
      }

      let requestOptions: CliRequestOptions;

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
          url: `${args.options.webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeFileUrl}')/CheckIn(comment=@a1,checkintype=@a2)?@a1='${formatting.encodeQueryParameter(args.options.publishMessage || '')}'&@a2=1`,
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

module.exports = new SpoPageSetCommand();